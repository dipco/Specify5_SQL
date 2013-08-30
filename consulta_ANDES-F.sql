USE [ANDES_F]
GO
/** Object: UserDefinedFunction [dbo].[ConsultaCompleta] Script Date: 08/13/2013 11:33:37 **/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author: Diego Perico Manrique
-- Create date: enero / 2013
/** Description: Esta función genera una consulta que permite recuperar todos los datos
almacenados en la base de datos de Specify 5 de la coleccion de Bacterias del
Museo de Historia Natural Andes Andes-F. Incluye los datos del ejemplar, determinaciones,
colecta y caracteristicas del ejemplar descritas en la base de datos.
Solo incluye los datos de la determinación considerada como actual. Crea los campos necesarios
para ingresar determinaciones anteriores, pero estos deben ser agregados luego desde otra
consulta.
**/
 
CREATE FUNCTION [dbo].[ConsultaCompleta]
()
RETURNS
@Completa TABLE
(
CollectionObjectID	INT,
--Accession
AccNumber	NVARCHAR (30),
AccStatus	NVARCHAR (30),
AccType	NVARCHAR (30),
DateAccessioned	INT,
DateRecived	INT,
AccRemarks	NTEXT,
--Catalog
CatalogNumber	FLOAT,
FieldNumber	VARCHAR (100),
AndesTnumber1	FLOAT,
DeterminadorText1	NVARCHAR (300),
CollectionObRemarks	NTEXT,
CatalogerFirstName	VARCHAR(150),
CatalogerLastName	VARCHAR (150),
CataloguedDateSPY	INT,
CataloguedDate	DATETIME NULL,
AltCatNumber	VARCHAR (100),
--Preparation
--1
PrepType1	VARCHAR (150),
FixationMethod1	VARCHAR (150),
[Count1]	INT,
PreparedFirstNameBy1	VARCHAR(150),
PreparedLastNameBy1	VARCHAR (150),
PreparedDateSPY1	INT,
PreparedDate1	DATETIME NULL,
--2
PrepType2	VARCHAR (150),
FixationMethod2	VARCHAR (150),
[Count2]	INT,
PreparedFirstNameBy2	VARCHAR(150),
PreparedLastNameBy2	VARCHAR (150),
PreparedDateSPY2	INT,
PreparedDate2	DATETIME NULL,
--3
PrepType3	VARCHAR (150),
FixationMethod3	VARCHAR (150),
[Count3]	INT,
PreparedFirstNameBy3	VARCHAR(150),
PreparedLastNameBy3	VARCHAR (150),
PreparedDateSPY3	INT,
PreparedDate3	DATETIME NULL,
--4
PrepType4	VARCHAR (150),
FixationMethod4	VARCHAR (150),
[Count4]	INT,
PreparedFirstNameBy4	VARCHAR(150),
PreparedLastNameBy4	VARCHAR (150),
PreparedDateSPY4	INT,
PreparedDate4	DATETIME NULL,
--Determination
TaxonID	INT,
--1
Kingdom1	VARCHAR (50),
KingdomAuthor1	VARCHAR (150),
--2
Division1	VARCHAR (50),
DivisionAuthor1	VARCHAR (150),
--3
Class1	VARCHAR (150),
ClassAuthor1	VARCHAR (150),
--4
[Order1]	VARCHAR (150),
OrderAuthor1	VARCHAR (150),
--5
--Suborder1 VARCHAR (150),
--SuborderAuthor1 VARCHAR (150),
--6
Family1	VARCHAR (150),
FamilyAuthor1	VARCHAR (150),
--7
Genus1	VARCHAR (150),
GenusAuthor1	VARCHAR (150),
--8
Species1	VARCHAR(150),
SpeciesAuthor1	VARCHAR (100),
--9
--Subspecies1 VARCHAR (150),
--SubspeciesAuthor1 VARCHAR (100),
DeterminerFirstName1	VARCHAR (100),
DeterminerLastName1	VARCHAR (100),
DeterminationDateSPY1	INT,
DeterminationDate1	DATETIME NULL,
TypeStatusName1	VARCHAR (100),
Current1	BIT,
DeterminationRemarks	NTEXT,
--Segunda Determinación
--1
Kingdom2	VARCHAR (50),
KingdomAuthor2	VARCHAR (150),
--2
Division2	VARCHAR (50),
DivisionAuthor2	VARCHAR (150),
--3
Class2	VARCHAR (150),
ClassAuthor2	VARCHAR (150),
--4
Order2	VARCHAR (150),
OrderAuthor2	VARCHAR (150),
--5
--Suborder2 VARCHAR (150),
--SuborderAuthor2 VARCHAR (150),
--6
Family2	VARCHAR (150),
FamilyAuthor2	VARCHAR (150),
--7
Genus2	VARCHAR (150),
GenusAuthor2	VARCHAR (150),
--8
Species2	VARCHAR(150),
SpeciesAuthor2	VARCHAR (100),
--9
--Subspecies2 VARCHAR (150),
--SubspeciesAuthor2 VARCHAR (100),
DeterminerFirstName2	VARCHAR (100),
DeterminerLastName2	VARCHAR (100),
DeterminationDateSPY2	INT,
DeterminationDate2	DATETIME NULL,
TypeStatusName2	VARCHAR (100),
Current2	BIT,
--Tercera Determinación
--1
--1
Kingdom3	VARCHAR (50),
KingdomAuthor3	VARCHAR (150),
--2
Division3	VARCHAR (50),
DivisionAuthor3	VARCHAR (150),
--3
Class3	VARCHAR (150),
ClassAuthor3	VARCHAR (150),
--4
Order3	VARCHAR (150),
OrderAuthor3	VARCHAR (150),
--5
Suborder3	VARCHAR (150),
SuborderAuthor3	VARCHAR (150),
--6
Family3	VARCHAR (150),
FamilyAuthor3	VARCHAR (150),
--7
Genus3	VARCHAR (150),
GenusAuthor3	VARCHAR (150),
--8
Species3	VARCHAR(150),
SpeciesAuthor3	VARCHAR (100),
--9
Subspecies3	VARCHAR (150),
SubspeciesAuthor3	VARCHAR (100),
DeterminerFirstName3	VARCHAR (100),
DeterminerLastName3	VARCHAR (100),
DeterminationDateSPY3	INT,
DeterminationDate3	DATETIME NULL,
TypeStatusName3	VARCHAR (100),
Current3	BIT,
--Collecting
Continent	VARCHAR (100),
Country	VARCHAR (100),
[State]	VARCHAR (100),
County	VARCHAR (100),
Locality	NTEXT,
Named	NTEXT,
MinElevation	REAL,
MaxElevation	REAL,
ElevationUnits	NVARCHAR (50),
ElevationAcuracy	REAL,
Latitude1	DECIMAL (12,10),
Longitude1	DECIMAL (13,10),
Latitude2	DECIMAL (12,10),
Longitude2	DECIMAL (13,10),
[Latitude/LongitudeAccuracy]	FLOAT,
[Latitude/LongitudeType]	VARCHAR (50),
LatLongMethod	VARCHAR (100),
ElevationMethod	VARCHAR (100),
StartDateSPY	INT,
StartDate	DATETIME NULL,
EndDateSPY	INT,
EndDate	DATETIME NULL,
StartTime	SMALLINT,
EndTime	SMALLINT,
Remarks	NTEXT,
AttractiveREMmethod	VARCHAR (150),
CollectionRemarksVerbaLocal	NTEXT,
--collector
CollectorFirstName1	VARCHAR (100),
CollectorLastName1	VARCHAR (100),
CollectorMiddle1	VARCHAR (100),
CollectorFirstName2	VARCHAR (100),
CollectorLastName2	VARCHAR (100),
CollectorMiddle2	VARCHAR (100),
CollectorFirstName3	VARCHAR (100),
CollectorLastName3	VARCHAR (100),
CollectorMiddle3	VARCHAR (100),
CollectorFirstName4	VARCHAR (100),
CollectorLastName4	VARCHAR (100),
CollectorMiddle4	VARCHAR (100),
--attributes
Stage	VARCHAR (150),
Sex	VARCHAR (150),
[Weight] FLOAT,
[%GrasaCondition]	NVARCHAR (50),
[%OcificaSnoutVentLenght]	FLOAT,
EnvergaduraHeight	FLOAT,
AnchGonaDerInsideHeighAper	FLOAT,
AnchGonaIzqInsideWidhtAper	FLOAT,
LenghtOfRightGonLengBill	FLOAT,
LenghtOfLeftGonLengGonad	FLOAT,
OvarioGranuladoBranchingAT	NVARCHAR (50),
LongGonOvMasGranLengMiddleToe	FLOAT,
ArticulosRelacionadosText4	NVARCHAR (50)
 
 
)
AS BEGIN
DECLARE @taxonomy TABLE
(
CollectionObject_ID	INT,
Field_Number	VARCHAR (100),
AndesTnumber_1	FLOAT,
DeterminadorText_1	NVARCHAR (300),
RemarksCollObj	TEXT,
Taxon_ID	INT,
--1
Kingdom_Name	VARCHAR (50),
Kingdom_Author	VARCHAR (150),
--2
Division_name	VARCHAR (150),
Division_Author	VARCHAR (150),
--3
Class_Name	VARCHAR (150),
Class_Author	VARCHAR (150),
--4
--SubClass VARCHAR (150),
--SubclassAuthor VARCHAR (150),
--5
Order_name	VARCHAR (150),
Order_Author	VARCHAR (150),
--6
--Suborder_name VARCHAR (150),
--Suborder_Author VARCHAR (150),
--7
Family_name	VARCHAR (150),
Family_Author	VARCHAR (150),
--8
Genus_Name	VARCHAR(150),
Genus_Author	VARCHAR (150),
--9
--Subgenus_Name VARCHAR(150),
--Subgenus_Author VARCHAR (150),
--10
Species_Name	VARCHAR(150),
Species_Author	VARCHAR (100),
--11
--Subspecies_Name VARCHAR (150),
--Subspecies_Author VARCHAR (100),
Determiner_FirstName	VARCHAR (100),
Determiner_LastName	VARCHAR (100),
[Determination_Date]	INT,
TypeStatus_Name	VARCHAR (100),
isCurrent	BIT,
Determination_Remarks	NTEXT
 
)
INSERT INTO @taxonomy
(
CollectionObject_ID,
Field_Number,
AndesTnumber_1,
DeterminadorText_1,
RemarksCollObj,
Taxon_ID,
--1,
Kingdom_Name,
Kingdom_Author,
--2
Division_name,
Division_Author,
--3
Class_Name,
Class_Author,
--4
Order_name,
Order_Author,
--5
--Suborder_name,
--Suborder_Author,
--6
Family_name,
Family_Author,
--7
Genus_Name,
Genus_Author,
--8
Species_Name,
Species_Author,
--9
--Subspecies_Name,
--Subspecies_Author,
 
--Determiner,
Determiner_FirstName,
Determiner_LastName,
[Determination_Date],
TypeStatus_Name,
isCurrent,
Determination_Remarks
 
)
SELECT
dbo.CollectionObject.CollectionObjectID,
dbo.CollectionObject.FieldNumber,
dbo.CollectionObject.Number1,
dbo.CollectionObject.Text1,
dbo.CollectionObject.Remarks,
dbo.TaxonName.TaxonNameID,
--1
--Kingdom
'Fungi',
--KingdomAuthor
'(L., 1753) R.T. Moore, 1980',
--2
--Division
CASE WHEN dbo.TaxonName.RankID = 30 THEN dbo.TaxonName.TaxonName
WHEN TaxonName1.RankID = 30 THEN TaxonName1.TaxonName
WHEN TaxonName2.RankID = 30 THEN TaxonName2.TaxonName
WHEN TaxonName3.RankID = 30 THEN TaxonName3.TaxonName
WHEN TaxonName4.RankID = 30 THEN TaxonName4.TaxonName
WHEN TaxonName5.RankID = 30 THEN TaxonName5.TaxonName
WHEN TaxonName6.RankID = 30 THEN TaxonName6.TaxonName
WHEN TaxonName7.RankID = 30 THEN TaxonName7.TaxonName
WHEN TaxonName8.RankID = 30 THEN TaxonName5.TaxonName
WHEN TaxonName9.RankID = 30 THEN TaxonName6.TaxonName
END
,
--DivisionAuthor
CASE WHEN dbo.TaxonName.RankID = 30 THEN dbo.TaxonName.Author
WHEN TaxonName1.RankID = 30 THEN TaxonName1.Author
WHEN TaxonName2.RankID = 30 THEN TaxonName2.Author
WHEN TaxonName3.RankID = 30 THEN TaxonName3.Author
WHEN TaxonName4.RankID = 30 THEN TaxonName4.Author
WHEN TaxonName5.RankID = 30 THEN TaxonName5.Author
WHEN TaxonName6.RankID = 30 THEN TaxonName6.Author
WHEN TaxonName7.RankID = 30 THEN TaxonName7.Author
WHEN TaxonName8.RankID = 30 THEN TaxonName8.Author
WHEN TaxonName9.RankID = 30 THEN TaxonName9.Author
END
,
--3
--Class
CASE WHEN dbo.TaxonName.RankID = 60 THEN dbo.TaxonName.TaxonName
WHEN TaxonName1.RankID = 60 THEN TaxonName1.TaxonName
WHEN TaxonName2.RankID = 60 THEN TaxonName2.TaxonName
WHEN TaxonName3.RankID = 60 THEN TaxonName3.TaxonName
WHEN TaxonName4.RankID = 60 THEN TaxonName4.TaxonName
WHEN TaxonName5.RankID = 60 THEN TaxonName5.TaxonName
WHEN TaxonName6.RankID = 60 THEN TaxonName6.TaxonName
WHEN TaxonName7.RankID = 60 THEN TaxonName7.TaxonName
WHEN TaxonName8.RankID = 60 THEN TaxonName5.TaxonName
WHEN TaxonName9.RankID = 60 THEN TaxonName6.TaxonName
END
,	
--ClassAuthor
CASE WHEN dbo.TaxonName.RankID = 60 THEN dbo.TaxonName.Author
WHEN TaxonName1.RankID = 60 THEN TaxonName1.Author
WHEN TaxonName2.RankID = 60 THEN TaxonName2.Author
WHEN TaxonName3.RankID = 60 THEN TaxonName3.Author
WHEN TaxonName4.RankID = 60 THEN TaxonName4.Author
WHEN TaxonName5.RankID = 60 THEN TaxonName5.Author
WHEN TaxonName6.RankID = 60 THEN TaxonName6.Author
WHEN TaxonName7.RankID = 60 THEN TaxonName7.Author
WHEN TaxonName8.RankID = 60 THEN TaxonName8.Author
WHEN TaxonName9.RankID = 60 THEN TaxonName9.Author
END
,
--4
--Order
CASE WHEN dbo.TaxonName.RankID = 100 THEN dbo.TaxonName.TaxonName
WHEN TaxonName1.RankID = 100 THEN TaxonName1.TaxonName
WHEN TaxonName2.RankID = 100 THEN TaxonName2.TaxonName
WHEN TaxonName3.RankID = 100 THEN TaxonName3.TaxonName
WHEN TaxonName4.RankID = 100 THEN TaxonName4.TaxonName
WHEN TaxonName5.RankID = 100 THEN TaxonName5.TaxonName
WHEN TaxonName6.RankID = 100 THEN TaxonName6.TaxonName
WHEN TaxonName7.RankID = 100 THEN TaxonName7.TaxonName
WHEN TaxonName8.RankID = 100 THEN TaxonName8.TaxonName
WHEN TaxonName9.RankID = 100 THEN TaxonName9.TaxonName
END
,
--OrderAuthor
CASE WHEN dbo.TaxonName.RankID = 100 THEN dbo.TaxonName.Author
WHEN TaxonName1.RankID = 100 THEN TaxonName1.Author
WHEN TaxonName2.RankID = 100 THEN TaxonName2.Author
WHEN TaxonName3.RankID = 100 THEN TaxonName3.Author
WHEN TaxonName4.RankID = 100 THEN TaxonName4.Author
WHEN TaxonName5.RankID = 100 THEN TaxonName5.Author
WHEN TaxonName6.RankID = 100 THEN TaxonName6.Author
WHEN TaxonName7.RankID = 100 THEN TaxonName7.Author
WHEN TaxonName8.RankID = 100 THEN TaxonName8.Author
WHEN TaxonName9.RankID = 100 THEN TaxonName9.Author
END
,
/* --5
--Suborder
CASE WHEN dbo.TaxonName.RankID = 110 THEN dbo.TaxonName.TaxonName
WHEN TaxonName1.RankID = 110 THEN TaxonName1.TaxonName
WHEN TaxonName2.RankID = 110 THEN TaxonName2.TaxonName
WHEN TaxonName3.RankID = 110 THEN TaxonName3.TaxonName
WHEN TaxonName4.RankID = 110 THEN TaxonName4.TaxonName
WHEN TaxonName5.RankID = 110 THEN TaxonName5.TaxonName
WHEN TaxonName6.RankID = 110 THEN TaxonName6.TaxonName
WHEN TaxonName7.RankID = 110 THEN TaxonName7.TaxonName
WHEN TaxonName8.RankID = 110 THEN TaxonName8.TaxonName
WHEN TaxonName9.RankID = 110 THEN TaxonName9.TaxonName
END
,
--SuborderAuthor
CASE WHEN dbo.TaxonName.RankID = 110 THEN dbo.TaxonName.Author
WHEN TaxonName1.RankID = 110 THEN TaxonName1.Author
WHEN TaxonName2.RankID = 110 THEN TaxonName2.Author
WHEN TaxonName3.RankID = 110 THEN TaxonName3.Author
WHEN TaxonName4.RankID = 110 THEN TaxonName4.Author
WHEN TaxonName5.RankID = 110 THEN TaxonName5.Author
WHEN TaxonName6.RankID = 110 THEN TaxonName6.Author
WHEN TaxonName7.RankID = 110 THEN TaxonName7.Author
WHEN TaxonName8.RankID = 110 THEN TaxonName8.Author
WHEN TaxonName9.RankID = 110 THEN TaxonName9.Author
END
, */
--6
--Family
CASE WHEN dbo.TaxonName.RankID = 140 THEN dbo.TaxonName.TaxonName
WHEN TaxonName1.RankID = 140 THEN TaxonName1.TaxonName
WHEN TaxonName2.RankID = 140 THEN TaxonName2.TaxonName
WHEN TaxonName3.RankID = 140 THEN TaxonName3.TaxonName
WHEN TaxonName4.RankID = 140 THEN TaxonName4.TaxonName
WHEN TaxonName5.RankID = 140 THEN TaxonName5.TaxonName
WHEN TaxonName6.RankID = 140 THEN TaxonName6.TaxonName
WHEN TaxonName7.RankID = 140 THEN TaxonName7.TaxonName
WHEN TaxonName8.RankID = 140 THEN TaxonName8.TaxonName
END
,
--FamilyAuthor
CASE WHEN dbo.TaxonName.RankID = 140 THEN dbo.TaxonName.Author
WHEN TaxonName1.RankID = 140 THEN TaxonName1.Author
WHEN TaxonName2.RankID = 140 THEN TaxonName2.Author
WHEN TaxonName3.RankID = 140 THEN TaxonName3.Author
WHEN TaxonName4.RankID = 140 THEN TaxonName4.Author
WHEN TaxonName5.RankID = 140 THEN TaxonName5.Author
WHEN TaxonName6.RankID = 140 THEN TaxonName6.Author
WHEN TaxonName7.RankID = 140 THEN TaxonName7.Author
WHEN TaxonName8.RankID = 140 THEN TaxonName8.Author
END
,
--7
--Genus
CASE WHEN dbo.TaxonName.RankID = 180 THEN dbo.TaxonName.TaxonName
WHEN TaxonName1.RankID = 180 THEN TaxonName1.TaxonName
WHEN TaxonName2.RankID = 180 THEN TaxonName2.TaxonName
WHEN TaxonName3.RankID = 180 THEN TaxonName3.TaxonName
WHEN TaxonName4.RankID = 180 THEN TaxonName4.TaxonName
WHEN TaxonName5.RankID = 180 THEN TaxonName5.TaxonName
END
,
--GenusAuthor
CASE WHEN dbo.TaxonName.RankID = 180 THEN dbo.TaxonName.Author
WHEN TaxonName1.RankID = 180 THEN TaxonName1.Author
WHEN TaxonName2.RankID = 180 THEN TaxonName2.Author
WHEN TaxonName3.RankID = 180 THEN TaxonName3.Author
WHEN TaxonName4.RankID = 180 THEN TaxonName4.Author
WHEN TaxonName5.RankID = 180 THEN TaxonName5.Author
END
,
--8
--Species
CASE WHEN dbo.TaxonName.RankID = 220 THEN dbo.TaxonName.TaxonName
WHEN TaxonName1.RankID = 220 THEN TaxonName1.TaxonName
WHEN TaxonName2.RankID = 220 THEN TaxonName2.TaxonName
END
,
--SpeciesAuthor
CASE WHEN dbo.TaxonName.RankID = 220 THEN dbo.TaxonName.Author
WHEN TaxonName1.RankID = 220 THEN TaxonName1.Author
WHEN TaxonName2.RankID = 220 THEN TaxonName2.Author
END
,
/*--9
--Subspecies
CASE WHEN dbo.TaxonName.RankID = 230 THEN dbo.TaxonName.TaxonName
WHEN taxonName1.RankID = 230 THEN dbo.TaxonName.TaxonName
END
,
--Subspecies Author
CASE WHEN dbo.TaxonName.RankID = 230 THEN dbo.TaxonName.Author
WHEN TaxonName1.RankID = 230 THEN dbo.TaxonName.Author
END
,*/
dbo.Agent.FirstName,
dbo.Agent.LastName,
dbo.Determination.[date],
dbo.Determination.TypeStatusName,
dbo.Determination.[Current],
dbo.Determination.Remarks
FROM
(
(
dbo.TaxonName
INNER JOIN dbo.Determination
ON dbo.TaxonName.TaxonNameID = dbo.Determination.TaxonNameID
)
INNER JOIN dbo.CollectionObject
ON dbo.Determination.BiologicalObjectID = dbo.CollectionObject.CollectionObjectID
) LEFT JOIN dbo.Agent
ON dbo.Determination.DeterminerID = dbo.Agent.AgentID
LEFT OUTER JOIN dbo.TaxonName as TaxonName1 on dbo.TaxonName.parentTaxonNameID = TaxonName1.TaxonNameID
LEFT OUTER JOIN dbo.TaxonName as TaxonName2 on TaxonName1.parentTaxonNameID = TaxonName2.TaxonNameID
LEFT OUTER JOIN dbo.TaxonName as TaxonName3 on TaxonName2.parentTaxonNameID = TaxonName3.TaxonNameID
LEFT OUTER JOIN dbo.TaxonName as TaxonName4 on TaxonName3.parentTaxonNameID = TaxonName4.TaxonNameID
LEFT OUTER JOIN dbo.TaxonName as TaxonName5 on TaxonName4.parentTaxonNameID = TaxonName5.TaxonNameID
LEFT OUTER JOIN dbo.TaxonName as TaxonName6 on TaxonName5.parentTaxonNameID = TaxonName6.TaxonNameID
LEFT OUTER JOIN dbo.TaxonName as TaxonName7 on TaxonName6.parentTaxonNameID = TaxonName7.TaxonNameID
LEFT OUTER JOIN dbo.TaxonName as TaxonName8 on TaxonName7.parentTaxonNameID = TaxonName8.TaxonNameID
LEFT OUTER JOIN dbo.TaxonName as TaxonName9 on TaxonName8.parentTaxonNameID = TaxonName9.TaxonNameID
LEFT OUTER JOIN dbo.TaxonName as TaxonName10 on TaxonName9.parentTaxonNameID = TaxonName10.TaxonNameID
WHERE dbo.determination.[current] = 1
DECLARE @Collecting TABLE
(
CollectionObject_ID1	INT,
Continent_Name	VARCHAR (100),
Country_Name	VARCHAR (100),
State_Name	VARCHAR (100),
County_name	VARCHAR (100),
Locality_Name	TEXT,
Named_Place	TEXT,
Min_Elevation	REAL,
Max_Elevation	REAL,
Elevation_Units	VARCHAR (50),
Elevation_Accuracy	REAL,
Latitude_1	DECIMAL (12,10),
Longitude_1	DECIMAL (13,10),
Latitude_2	DECIMAL (12,10),
Longitude_2	DECIMAL (13, 10),
[Latitude/Longitude_Accuracy]	FLOAT,
[Latitude/Longitude_Type]	VARCHAR (50),
LatLong_Method	VARCHAR (100),
Elevation_Method	VARCHAR (100),
Star_Date	INT,
End_Date	INT,
Start_Time	SMALLINT,
End_Time	SMALLINT,
CollectionRemarks_VerbatimLoca	TEXT,
AttractiveREM_method	VARCHAR (150),
Remarks	TEXT
)
INSERT INTO @Collecting
(
CollectionObject_ID1,
Continent_Name,
Country_Name,
State_Name,
County_name,
Locality_Name,
Named_Place,
Min_Elevation,
Max_Elevation,
Elevation_Units,
Elevation_Accuracy,
Latitude_1,
Longitude_1,
Latitude_2,
Longitude_2,
[Latitude/Longitude_Accuracy],
[Latitude/Longitude_Type],
LatLong_Method,
Elevation_Method,
Star_Date,
End_Date,
Start_Time,
End_Time,
CollectionRemarks_VerbatimLoca,
AttractiveREM_method,
Remarks
 
 
)
SELECT
dbo.CollectionObject.CollectionObjectID,
dbo.[Geography].ContinentOrOcean,
dbo.[Geography].Country,
dbo.[Geography].[State],
dbo.[Geography].County,
dbo.Locality.LocalityName,
dbo.Locality.NamedPlace,
dbo.Locality.MinElevation,
dbo.Locality.MaxElevation,
dbo.Locality.OriginalElevationUnit,
dbo.Locality.ElevationAccuracy,
dbo.Locality.Latitude1,
dbo.Locality.Longitude1,
dbo.Locality.Latitude2,
dbo.Locality.Longitude2,
dbo.Locality.OriginalLatLongUnit,
dbo.Locality.LatLongType,
dbo.Locality.LatLongMethod,
dbo.Locality.ElevationMethod,
dbo.CollectingEvent.StartDate,
dbo.CollectingEvent.EndDate,
dbo.CollectingEvent.StartTime,
dbo.CollectingEvent.EndTime,
dbo.CollectingEvent.VerbatimLocality,
dbo.CollectingEvent.Method,
dbo.CollectingEvent.Remarks
FROM
(
(
dbo.CollectingEvent
LEFT JOIN dbo.Locality
ON dbo.CollectingEvent.LocalityID = dbo.Locality.LocalityID
)
LEFT JOIN dbo.[Geography]
ON dbo.Locality.GeographyID = dbo.[Geography].GeographyID
)
RIGHT JOIN dbo.CollectionObject
ON dbo.CollectingEvent.CollectingEventID = dbo.CollectionObject.CollectingEventID
 
DECLARE @collector1 TABLE
(
CollectionObjectC1_ID	INT,
Collector1FirstName	VARCHAR (100),
Collector1LastName	VARCHAR (100),
Collector1Middle	VARCHAR (100)
)
INSERT INTO @collector1
(
CollectionObjectC1_ID,
Collector1FirstName,
Collector1LastName,
Collector1Middle
)
SELECT
 
dbo.CollectionObject.CollectionObjectID,
CASE WHEN dbo.Collectors.[Order] = 1 THEN dbo.Agent.FirstName
ELSE NULL
END
,
CASE WHEN dbo.Collectors.[Order] = 1 THEN dbo.Agent.LastName
ELSE NULL	
END
,
CASE WHEN dbo.Collectors.[Order] = 1 THEN dbo.Agent.MiddleInitial
ELSE NULL
END
 
FROM	
 
(
(
dbo.CollectionObject
INNER JOIN dbo.CollectingEvent
ON dbo.CollectionObject.CollectingEventID = dbo.CollectingEvent.CollectingEventID
)
INNER JOIN dbo.Collectors
ON dbo.CollectingEvent.CollectingEventID = dbo.Collectors.CollectingEventID
)
INNER JOIN dbo.Agent
ON dbo.Collectors.AgentID = dbo.Agent.AgentID
WHERE (((dbo.Collectors.[Order])= 1))
 
DECLARE @collector2 TABLE
(
CollectionObjectC2_ID	INT,
Collector2FirstName	VARCHAR (100),
Collector2LastName	VARCHAR (100),
Collector2Middle	VARCHAR (100)
)
INSERT INTO @collector2
(
CollectionObjectC2_ID,
Collector2FirstName,
Collector2LastName,
Collector2Middle
)
SELECT
 
dbo.CollectionObject.CollectionObjectID,
CASE WHEN dbo.Collectors.[Order] = 2 THEN dbo.Agent.FirstName
ELSE NULL
END
,
CASE WHEN dbo.Collectors.[Order] = 2 THEN dbo.Agent.LastName
ELSE NULL	
END
,
CASE WHEN dbo.Collectors.[Order] = 2 THEN dbo.Agent.MiddleInitial
ELSE NULL
END
 
FROM	
 
(
(
dbo.CollectionObject
INNER JOIN dbo.CollectingEvent
ON dbo.CollectionObject.CollectingEventID = dbo.CollectingEvent.CollectingEventID
)
INNER JOIN dbo.Collectors
ON dbo.CollectingEvent.CollectingEventID = dbo.Collectors.CollectingEventID
)
INNER JOIN dbo.Agent
ON dbo.Collectors.AgentID = dbo.Agent.AgentID
WHERE (((dbo.Collectors.[Order])= 2))
 
DECLARE @collector3 TABLE
(
CollectionObjectC3_ID	INT,
Collector3FirstName	VARCHAR (100),
Collector3LastName	VARCHAR (100),
Collector3Middle	VARCHAR (100)
)
INSERT INTO @collector3
(
CollectionObjectC3_ID,
Collector3FirstName,
Collector3LastName,
Collector3Middle
)
SELECT
 
dbo.CollectionObject.CollectionObjectID,
CASE WHEN dbo.Collectors.[Order] = 3 THEN dbo.Agent.FirstName
ELSE NULL
END
,
CASE WHEN dbo.Collectors.[Order] = 3 THEN dbo.Agent.LastName
ELSE NULL	
END
,
CASE WHEN dbo.Collectors.[Order] = 3 THEN dbo.Agent.MiddleInitial
ELSE NULL
END
 
FROM	
 
(
(
dbo.CollectionObject
INNER JOIN dbo.CollectingEvent
ON dbo.CollectionObject.CollectingEventID = dbo.CollectingEvent.CollectingEventID
)
INNER JOIN dbo.Collectors
ON dbo.CollectingEvent.CollectingEventID = dbo.Collectors.CollectingEventID
)
INNER JOIN dbo.Agent
ON dbo.Collectors.AgentID = dbo.Agent.AgentID
WHERE (((dbo.Collectors.[Order])= 3))
 
DECLARE @collector4 TABLE
(
CollectionObjectC4_ID	INT,
Collector4FirstName	VARCHAR (100),
Collector4LastName	VARCHAR (100),
Collector4Middle	VARCHAR (100)
)
INSERT INTO @collector4
(
CollectionObjectC4_ID,
Collector4FirstName,
Collector4LastName,
Collector4Middle
)
SELECT
 
dbo.CollectionObject.CollectionObjectID,
CASE WHEN dbo.Collectors.[Order] = 4 THEN dbo.Agent.FirstName
ELSE NULL
END
,
CASE WHEN dbo.Collectors.[Order] = 4 THEN dbo.Agent.LastName
ELSE NULL	
END
,
CASE WHEN dbo.Collectors.[Order] = 4 THEN dbo.Agent.MiddleInitial
ELSE NULL
END
 
FROM	
 
(
(
dbo.CollectionObject
INNER JOIN dbo.CollectingEvent
ON dbo.CollectionObject.CollectingEventID = dbo.CollectingEvent.CollectingEventID
)
INNER JOIN dbo.Collectors
ON dbo.CollectingEvent.CollectingEventID = dbo.Collectors.CollectingEventID
)
INNER JOIN dbo.Agent
ON dbo.Collectors.AgentID = dbo.Agent.AgentID
WHERE (((dbo.Collectors.[Order])= 4))
 
DECLARE @collectors TABLE
(
CollectionObject_ID2	INT,
CollectorFirstName_1	VARCHAR (100),
CollectorLastName_1	VARCHAR (100),
CollectorMiddle_1	VARCHAR (100),
CollectorFirstName_2	VARCHAR (100),
CollectorLastName_2	VARCHAR (100),
CollectorMiddle_2	VARCHAR (100),
CollectorFirstName_3	VARCHAR (100),
CollectorLastName_3	VARCHAR (100),
CollectorMiddle_3	VARCHAR (100),
CollectorFirstName_4	VARCHAR (100),
CollectorLastName_4	VARCHAR (100),
CollectorMiddle_4	VARCHAR (100)
)
INSERT INTO @collectors
(
CollectionObject_ID2,
CollectorFirstName_1,
CollectorLastName_1,
CollectorMiddle_1,
CollectorFirstName_2,
CollectorLastName_2,
CollectorMiddle_2,
CollectorFirstName_3,
CollectorLastName_3,
CollectorMiddle_3,
CollectorFirstName_4,
CollectorLastName_4,
CollectorMiddle_4
)
 
SELECT
 
CollectionObjectC1_ID,
Collector1FirstName,
Collector1LastName,
Collector1Middle,
Collector2FirstName,
Collector2LastName,
Collector2Middle,
Collector3FirstName,
Collector3LastName,
Collector3Middle,
Collector4FirstName,
Collector4LastName,
Collector4Middle
 
FROM	
 
@Collector1
LEFT JOIN @Collector2
ON CollectionObjectC1_ID = CollectionObjectC2_ID
LEFT JOIN @Collector3
ON CollectionObjectC1_ID = CollectionObjectC3_ID
LEFT JOIN @Collector4
ON CollectionObjectC1_ID = CollectionObjectC4_ID
 
DECLARE @Attributes TABLE
(
CollectionObject_ID3	INT,
att_Stage	VARCHAR (150),
att_Sex	VARCHAR (150),
att_Weight FLOAT,
[att_%GrasaCondition]	NVARCHAR (50),
[att_%OcificaSnoutVentLenght]	FLOAT,
att_EnvergaduraHeight	FLOAT,
att_AnchGonaDerInsideHeighAper	FLOAT,
att_AnchGonaIzqInsideWidhtAper	FLOAT,
att_LenghtOfRightGonLengBill	FLOAT,
att_LenghtOfLeftGonLengGonad	FLOAT,
att_OvarioGranuladoBranchingAT	NVARCHAR (50),
att_LongGonOvMasGranLengMiddleToe	FLOAT,
att_ArticulosRelacionadosText4	NVARCHAR (50)
 
)
 
INSERT INTO @Attributes
(
CollectionObject_ID3,
att_Stage,
att_Sex,
att_Weight,
[att_%GrasaCondition],
[att_%OcificaSnoutVentLenght],
att_EnvergaduraHeight,
att_AnchGonaDerInsideHeighAper,
att_AnchGonaIzqInsideWidhtAper,
att_LenghtOfRightGonLengBill,
att_LenghtOfLeftGonLengGonad,
att_OvarioGranuladoBranchingAT,
att_LongGonOvMasGranLengMiddleToe,
att_ArticulosRelacionadosText4
)
SELECT
 
dbo.BiologicalObjectAttributes.BiologicalObjectAttributesID,
dbo.BiologicalObjectAttributes.Stage,
dbo.BiologicalObjectAttributes.Sex,
dbo.BiologicalObjectAttributes.[Weight],
dbo.BiologicalObjectAttributes.Condition,
dbo.BiologicalObjectAttributes.SnoutVentLength,
dbo.BiologicalObjectAttributes.Height,
dbo.BiologicalObjectAttributes.HeightFinalWhorl,
dbo.BiologicalObjectAttributes.InsideHeightAperture,
dbo.BiologicalObjectAttributes.LengthBill,
dbo.BiologicalObjectAttributes.LengthGonad,
dbo.BiologicalObjectAttributes.BranchingAt,
dbo.BiologicalObjectAttributes.LengthMiddleToe,
dbo.BiologicalObjectAttributes.Text4
 
from dbo.BiologicalObjectAttributes
 
DECLARE @PrimPrep TABLE
(
PrimPrepMethodID	INT,
CatalogNumberPrep	FLOAT
)
INSERT INTO @PrimPrep
(
PrimPrepMethodID,
CatalogNumberPrep	
)
SELECT
MIN (dbo.CollectionObject.CollectionObjectID),
dbo.CollectionObjectCatalog.CatalogNumber
FROM
CollectionObjectCatalog
INNER JOIN CollectionObject
ON CollectionObjectCatalog.CollectionObjectCatalogID = CollectionObject.DerivedFromID
GROUP BY CatalogNumber
 
DECLARE @catalog TABLE
(
CollectionObject_ID4	INT,
PrepMethodID	INT,
Catalog_Number	FLOAT,
Preparation_Method	VARCHAR (150),
Fixation_Method	VARCHAR (150),
CatalogerFirst_Name	VARCHAR(150),
CatalogerLast_Name	VARCHAR (150),
Catalogued_Date	INT,
Count_1	INT,
PreviousNumber VARCHAR (100)
)
INSERT INTO @catalog
(
CollectionObject_ID4,
PrepMethodID,
Catalog_Number,
Preparation_Method,
Fixation_Method,
CatalogerFirst_Name,
CatalogerLast_Name,
Catalogued_Date,
Count_1,
PreviousNumber
)
 
SELECT
dbo.CollectionObjectCatalog.CollectionObjectCatalogID,
 
CollectionObject_1.CollectionObjectID,
 
dbo.CollectionObjectCatalog.CatalogNumber,
 
CASE WHEN CollectionObject_1.PreparationMethod IS NULL THEN 'Dry'
ELSE CollectionObject_1.PreparationMethod
END
-- AS PreparationMethod
,
CollectionObject_1.[description],
dbo.Agent.FirstName,
dbo.Agent.LastName,
dbo.CollectionObjectCatalog.CatalogedDate,
CollectionObject_1.[Count],
dbo.CollectionObjectCatalog.Name
FROM
(
dbo.CollectionObject
AS CollectionObject_1
RIGHT JOIN
(
dbo.Determination
INNER JOIN
(
dbo.CollectionObjectCatalog
INNER JOIN dbo.CollectionObject
ON dbo.CollectionObjectCatalog.CollectionObjectCatalogID = dbo.CollectionObject.CollectionObjectID
)
 
ON dbo.Determination.BiologicalObjectID = dbo.CollectionObjectCatalog.CollectionObjectCatalogID
)
ON CollectionObject_1.DerivedFromID = dbo.CollectionObject.CollectionObjectID
)
LEFT JOIN dbo.Agent
ON dbo.CollectionObjectCatalog.CatalogerID = dbo.Agent.AgentID
GROUP BY
dbo.CollectionObjectCatalog.CollectionObjectCatalogID,
dbo.CollectionObjectCatalog.CatalogNumber,
CollectionObject_1.PreparationMethod,
CollectionObject_1.[description],
CollectionObject_1.CollectionObjectID,
CollectionObject_1.[Count],
dbo.Agent.FirstName,
dbo.Agent.LastName,
dbo.CollectionObjectCatalog.CatalogedDate,
dbo.CollectionObjectCatalog.Name
 
DECLARE @Prepa TABLE
(
CollectionObject_ID5	INT,
Prepared_Date	INT,
PreparedBy_FirstName	VARCHAR (150),
PreparedBy_LastName	VARCHAR (150)
)
INSERT INTO @Prepa
(
CollectionObject_ID5,
Prepared_Date,
PreparedBy_FirstName,
PreparedBy_LastName
)
SELECT
dbo.Preparation.PreparationID,
dbo.Preparation.PreparedDate,
dbo.Agent.FirstName,
dbo.Agent.LastName
FROM dbo.Preparation
INNER JOIN dbo.Agent ON PreparedByID = AgentID
DECLARE @Access TABLE
(
CollectionObject_ID6	INT,
Acc_Number	NVARCHAR (30),
Acc_Status	NVARCHAR (30),
Acc_Type	NVARCHAR (30),
Date_Accessioned	INT,
Date_Recived	INT,
Acc_Remarks	NTEXT
)
INSERT INTO @Access
(
CollectionObject_ID6,
Acc_Number,
Acc_Status,
Acc_Type,
Date_Accessioned,
Date_Recived,
Acc_Remarks
)
SELECT
dbo.CollectionObjectCatalog.CollectionObjectCatalogID,
dbo.Accession.Number,
dbo.Accession.[Status],
dbo.Accession.[Type],
dbo.Accession.DateAccessioned,
dbo.Accession.DateReceived,
dbo.Accession.Remarks
 
from dbo.Accession
INNER JOIN
dbo.CollectionObjectCatalog
ON dbo.Accession.AccessionID = dbo.CollectionObjectCatalog.AccessionID
 
INSERT INTO @Completa
(
CollectionObjectID,
--Accession
AccNumber,
AccStatus,
AccType,
DateAccessioned,
DateRecived,
AccRemarks,
 
--Catalog
CatalogNumber,
FieldNumber,
AndesTnumber1,
DeterminadorText1,
CollectionObRemarks,
CatalogerFirstName,
CatalogerLastName,
CataloguedDateSPY,
AltCatNumber,
 
--Preparation,
--1,
PrepType1,
FixationMethod1,
[Count1],
PreparedFirstNameBy1,
PreparedLastNameBy1,
PreparedDateSPY1,
--Determination,
TaxonID,
--1,
Kingdom1,
KingdomAuthor1,
--2
Division1,
DivisionAuthor1,
--3
Class1,
ClassAuthor1,
--4
Order1,
OrderAuthor1,
--5
-- Suborder1,
-- SuborderAuthor1,
--6
Family1,
FamilyAuthor1,
--7
Genus1,
GenusAuthor1,
--8
Species1,
SpeciesAuthor1,
--9
-- Subspecies1,
-- SubspeciesAuthor1,
 
--Determiner,
DeterminerFirstName1,
DeterminerLastName1,
DeterminationDateSPY1,
TypeStatusName1,
Current1,
DeterminationRemarks,
--Collecting,
Continent,
Country,
[State],
County,
Locality,
Named,
MinElevation,
MaxElevation,
ElevationUnits,
ElevationAcuracy,
Latitude1,
Longitude1,
Latitude2,
Longitude2,
[Latitude/LongitudeAccuracy],
[Latitude/LongitudeType],
LatLongMethod,
ElevationMethod,
StartDateSPY,
EndDateSPY,
StartTime,
EndTime,
Remarks,
AttractiveREMmethod,
CollectionRemarksVerbaLocal,
--collector,
CollectorFirstName1,
CollectorLastName1,
CollectorMiddle1,
CollectorFirstName2,
CollectorLastName2,
CollectorMiddle2,
CollectorFirstName3,
CollectorLastName3,
CollectorMiddle3,
CollectorFirstName4,
CollectorLastName4,
CollectorMiddle4,
--attributes,
Stage,
Sex,
[Weight],
[%GrasaCondition],
[%OcificaSnoutVentLenght],
EnvergaduraHeight,
AnchGonaDerInsideHeighAper,
AnchGonaIzqInsideWidhtAper,
LenghtOfRightGonLengBill,
LenghtOfLeftGonLengGonad,
OvarioGranuladoBranchingAT,
LongGonOvMasGranLengMiddleToe,
ArticulosRelacionadosText4
 
)
SELECT
 
CollectionObject_ID,
--Accession
Acc_Number,
Acc_Status,
Acc_Type,
Date_Accessioned,
Date_Recived,
Acc_Remarks,
--Catalog
Catalog_Number,
Field_Number,
AndesTnumber_1,
DeterminadorText_1,
RemarksCollObj,
'Coleccion',
'Hongos',
Catalogued_Date,
PreviousNumber,
--Preparation
--1
Preparation_Method,
Fixation_Method,
Count_1,
PreparedBy_FirstName,
PreparedBy_LastName,
Prepared_Date,
--Determination
Taxon_ID,
--1,
Kingdom_Name,
Kingdom_Author,
--2
Division_name,
Division_Author,
--3
CASE WHEN Class_Name IS NULL THEN 'NOASIGNADO'
ELSE Class_Name
END,
Class_Author,
--4
Order_name,
Order_Author,
--5
--Suborder_name,
--Suborder_Author,
--6
Family_name,
Family_Author,
--7
Genus_Name,
Genus_Author,
--8
Species_Name,
Species_Author,
--9
--Subspecies_Name,
--Subspecies_Author,
--Determiner,
Determiner_FirstName,
Determiner_LastName,
[Determination_Date],
TypeStatus_Name,
isCurrent,
Determination_Remarks,
--Collecting
Continent_Name,
Country_Name,
State_Name,
County_name,
Locality_Name,
Named_Place,
Min_Elevation,
Max_Elevation,
Elevation_Units,
Elevation_Accuracy,
Latitude_1,
Longitude_1,
Latitude_2,
Longitude_2,
[Latitude/Longitude_Accuracy],
[Latitude/Longitude_Type],
LatLong_Method,
Elevation_Method,
Star_Date,
End_Date,
Start_Time,
End_Time,
Remarks,
AttractiveREM_method,
CollectionRemarks_VerbatimLoca,
--collector
CollectorFirstName_1,
CollectorLastName_1,
CollectorMiddle_1,
CollectorFirstName_2,
CollectorLastName_2,
CollectorMiddle_2,
CollectorFirstName_3,
CollectorLastName_3,
CollectorMiddle_3,
CollectorFirstName_4,
CollectorLastName_4,
CollectorMiddle_4,
--attributes
att_Stage,
att_Sex,
att_Weight,
[att_%GrasaCondition],
[att_%OcificaSnoutVentLenght],
att_EnvergaduraHeight,
att_AnchGonaDerInsideHeighAper,
att_AnchGonaIzqInsideWidhtAper,
att_LenghtOfRightGonLengBill,
att_LenghtOfLeftGonLengGonad,
att_OvarioGranuladoBranchingAT,
att_LongGonOvMasGranLengMiddleToe,
att_ArticulosRelacionadosText4
 
FROM @Taxonomy
INNER JOIN @Collecting
ON CollectionObject_ID = CollectionObject_ID1
LEFT JOIN @collectors
ON CollectionObject_ID = CollectionObject_ID2
LEFT JOIN @Attributes
ON CollectionObject_ID = CollectionObject_ID3
LEFT JOIN (@catalog
INNER JOIN @PrimPrep on PrepMethodID = PrimPrepMethodID )
ON CollectionObject_ID = CollectionObject_ID4
LEFT JOIN @PREPA
on CollectionObject_ID = CollectionObject_ID5
LEFT JOIN @Access
ON CollectionObject_ID = CollectionObject_ID6
RETURN
END

OtDetANDES-F.sql
SQL

USE [ANDES_F]
GO
/****** Object: UserDefinedFunction [dbo].[OtrasDeterminaciones] Script Date: 08/13/2013 12:05:51 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author: Diego Perico M.
-- Create date: 20/Jan/2013
/** Description: Esta funcion genera una consulta sobre la base de datos de
Specify 5 que regresa la información de determinaciones de
los ejemplares diferentes a la determinación actual.
Determinaciones historicas que luego deben ser integradas a
la consulta principal mediante una consulta de datos anexados
en los campos correspondientes.
**/
 
CREATE FUNCTION [dbo].[OtrasDeterminaciones]
()
 
RETURNS
@OtrasDet TABLE
(
CollectionObjectID INT,
Kingdom1 VARCHAR (10),
KingdomAuthor1 VARCHAR (50),
Class1 VARCHAR (50),
ClassAuthor1 VARCHAR (50),
[Order1] VARCHAR (50),
OrderAuthor1 VARCHAR (50),
Family1	VARCHAR (50),
FamilyAuthor1 VARCHAR (50),
Subfamily1	VARCHAR (50),
SubfamilyAuthor1 VARCHAR (50),
Subtribu1	VARCHAR (50),
SubtribuAuthor1 VARCHAR (50),
Genus1	VARCHAR (50),
GenusAuthor1 VARCHAR (50),
Subgenus1 VARCHAR (50),
SubgenusAuthor1 VARCHAR (50),
Section1 VARCHAR (50),
SectionAuthor1 VARCHAR (50),
Species1	VARCHAR(50),
SpeciesAuthor1 VARCHAR (100),
Subspecies1 VARCHAR (50),
SubspeciesAuthor1 VARCHAR (100),
Variety1 VARCHAR (50),
VarietyAuthor1 VARCHAR (50),
DeterminerFirstName1 VARCHAR (100),
DeterminerLastName1 VARCHAR (100),
[DeterminationDate1] VARCHAR (10),
TypeStatusName1 VARCHAR (100),
DeterminationID INT
)
AS BEGIN	
DECLARE @taxonomy TABLE
(
CollectionObject_ID INT,
-- Field_Number VARCHAR (20),
-- Andes_T VARCHAR (20),
-- Taxon_ID INT,
Kingdom_Name VARCHAR (10),
Kingdom_Author VARCHAR (50),
Class_name VARCHAR (50),
Class_Author VARCHAR (50),
Order_name VARCHAR (50),
Order_Author VARCHAR (50),
Family_name VARCHAR (50),
Family_Author VARCHAR (50),
SubFamily_name VARCHAR (50),
SubFamily_Author VARCHAR (50),
Subtribu_Name VARCHAR (50),
Subtribu_Author VARCHAR (50),
Genus_Name	VARCHAR(50),
Genus_Author VARCHAR (50),
Subgenus_Name	VARCHAR(50),
Subgenus_Author VARCHAR (50),
Section_Name VARCHAR (50),
Section_Author VARCHAR (50),
Species_Name VARCHAR(50),
Species_Author VARCHAR (100),
Subspecies_Name VARCHAR (50),
Subspecies_Author VARCHAR (100),
Variety_Name VARCHAR (50),
Variety_Author VARCHAR (100),
Determiner_FirstName VARCHAR (100),
Determiner_LastName VARCHAR (100),
[Determination_Date] VARCHAR (10),
TypeStatus_Name VARCHAR (100),
Determination_ID INT
 
)
INSERT INTO @taxonomy
(
CollectionObject_ID,
-- Field_Number,
-- Andes_T,
-- Taxon_ID,
Kingdom_Name,
Kingdom_Author,
Class_Name,
Class_Author,
Order_Name,
Order_Author,
Family_Name,
Family_Author,
SubFamily_Name,
SubFamily_Author,
Subtribu_Name,
Subtribu_Author,
Genus_Name,
Genus_Author,
Subgenus_Name,
Subgenus_Author,
Section_Name,
Section_Author,
Species_Name,
Species_Author,
Subspecies_Name,
Subspecies_Author,
Variety_Name,
Variety_Author,
Determiner_FirstName,
Determiner_LastName,
Determination_Date,
TypeStatus_Name,
Determination_ID
)
SELECT
dbo.CollectionObject.CollectionObjectID,
-- dbo.CollectionObject.FieldNumber,
-- dbo.CollectionObject.Number1,
-- dbo.TaxonName.TaxonNameID,
--Kingdom
'Plantae',
--Kingdom Author
'Haeckel, 1866',
--class
CASE WHEN dbo.TaxonName.RankID = 60 THEN dbo.TaxonName.TaxonName
WHEN TaxonName1.RankID = 60 THEN TaxonName1.TaxonName
WHEN TaxonName2.RankID = 60 THEN TaxonName2.TaxonName
WHEN TaxonName3.RankID = 60 THEN TaxonName3.TaxonName
WHEN TaxonName4.RankID = 60 THEN TaxonName4.TaxonName
WHEN TaxonName5.RankID = 60 THEN TaxonName5.TaxonName
WHEN TaxonName6.RankID = 60 THEN TaxonName6.TaxonName
WHEN TaxonName7.RankID = 60 THEN TaxonName7.TaxonName
WHEN TaxonName8.RankID = 60 THEN TaxonName5.TaxonName
WHEN TaxonName9.RankID = 60 THEN TaxonName6.TaxonName
WHEN TaxonName10.RankID = 60 THEN TaxonName7.TaxonName
END
,
--ClassAuthor
CASE WHEN dbo.TaxonName.RankID = 60 THEN dbo.TaxonName.Author
WHEN TaxonName1.RankID = 60 THEN TaxonName1.Author
WHEN TaxonName2.RankID = 60 THEN TaxonName2.Author
WHEN TaxonName3.RankID = 60 THEN TaxonName3.Author
WHEN TaxonName4.RankID = 60 THEN TaxonName4.Author
WHEN TaxonName5.RankID = 60 THEN TaxonName5.Author
WHEN TaxonName6.RankID = 60 THEN TaxonName6.Author
WHEN TaxonName7.RankID = 60 THEN TaxonName7.Author
WHEN TaxonName8.RankID = 60 THEN TaxonName8.Author
WHEN TaxonName9.RankID = 60 THEN TaxonName9.Author
WHEN TaxonName10.RankID = 60 THEN TaxonName10.Author
END
,
--Order
CASE WHEN dbo.TaxonName.RankID = 100 THEN dbo.TaxonName.TaxonName
WHEN TaxonName1.RankID = 100 THEN TaxonName1.TaxonName
WHEN TaxonName2.RankID = 100 THEN TaxonName2.TaxonName
WHEN TaxonName3.RankID = 100 THEN TaxonName3.TaxonName
WHEN TaxonName4.RankID = 100 THEN TaxonName4.TaxonName
WHEN TaxonName5.RankID = 100 THEN TaxonName5.TaxonName
WHEN TaxonName6.RankID = 100 THEN TaxonName6.TaxonName
WHEN TaxonName7.RankID = 100 THEN TaxonName7.TaxonName
WHEN TaxonName8.RankID = 100 THEN TaxonName8.TaxonName
WHEN TaxonName9.RankID = 100 THEN TaxonName9.TaxonName
END
,
--OrderAuthor
CASE WHEN dbo.TaxonName.RankID = 100 THEN dbo.TaxonName.Author
WHEN TaxonName1.RankID = 100 THEN TaxonName1.Author
WHEN TaxonName2.RankID = 100 THEN TaxonName2.Author
WHEN TaxonName3.RankID = 100 THEN TaxonName3.Author
WHEN TaxonName4.RankID = 100 THEN TaxonName4.Author
WHEN TaxonName5.RankID = 100 THEN TaxonName5.Author
WHEN TaxonName6.RankID = 100 THEN TaxonName6.Author
WHEN TaxonName7.RankID = 100 THEN TaxonName7.Author
WHEN TaxonName8.RankID = 100 THEN TaxonName8.Author
WHEN TaxonName9.RankID = 100 THEN TaxonName9.Author
END
,
--Family
CASE WHEN dbo.TaxonName.RankID = 140 THEN dbo.TaxonName.TaxonName
WHEN TaxonName1.RankID = 140 THEN TaxonName1.TaxonName
WHEN TaxonName2.RankID = 140 THEN TaxonName2.TaxonName
WHEN TaxonName3.RankID = 140 THEN TaxonName3.TaxonName
WHEN TaxonName4.RankID = 140 THEN TaxonName4.TaxonName
WHEN TaxonName5.RankID = 140 THEN TaxonName5.TaxonName
WHEN TaxonName6.RankID = 140 THEN TaxonName6.TaxonName
WHEN TaxonName7.RankID = 140 THEN TaxonName7.TaxonName
WHEN TaxonName8.RankID = 140 THEN TaxonName8.TaxonName
END
,
--FamilyAuthor
CASE WHEN dbo.TaxonName.RankID = 140 THEN dbo.TaxonName.Author
WHEN TaxonName1.RankID = 140 THEN TaxonName1.Author
WHEN TaxonName2.RankID = 140 THEN TaxonName2.Author
WHEN TaxonName3.RankID = 140 THEN TaxonName3.Author
WHEN TaxonName4.RankID = 140 THEN TaxonName4.Author
WHEN TaxonName5.RankID = 140 THEN TaxonName5.Author
WHEN TaxonName6.RankID = 140 THEN TaxonName6.Author
WHEN TaxonName7.RankID = 140 THEN TaxonName7.Author
WHEN TaxonName8.RankID = 140 THEN TaxonName8.Author
END
,
--SubFamily
CASE WHEN dbo.TaxonName.RankID = 150 THEN dbo.TaxonName.TaxonName
WHEN TaxonName1.RankID = 150 THEN TaxonName1.TaxonName
WHEN TaxonName2.RankID = 150 THEN TaxonName2.TaxonName
WHEN TaxonName3.RankID = 150 THEN TaxonName3.TaxonName
WHEN TaxonName4.RankID = 150 THEN TaxonName4.TaxonName
WHEN TaxonName5.RankID = 150 THEN TaxonName5.TaxonName
WHEN TaxonName6.RankID = 150 THEN TaxonName6.TaxonName
WHEN TaxonName7.RankID = 150 THEN TaxonName7.TaxonName
END
,
--SubFamilyAuthor
CASE WHEN dbo.TaxonName.RankID = 150 THEN dbo.TaxonName.Author
WHEN TaxonName1.RankID = 150 THEN TaxonName1.Author
WHEN TaxonName2.RankID = 150 THEN TaxonName2.Author
WHEN TaxonName3.RankID = 150 THEN TaxonName3.Author
WHEN TaxonName4.RankID = 150 THEN TaxonName4.Author
WHEN TaxonName5.RankID = 150 THEN TaxonName5.Author
WHEN TaxonName6.RankID = 150 THEN TaxonName6.Author
WHEN TaxonName7.RankID = 150 THEN TaxonName7.Author
END
,
--SubTribu
CASE WHEN dbo.TaxonName.RankID = 170 THEN dbo.TaxonName.TaxonName
WHEN TaxonName1.RankID = 170 THEN TaxonName1.TaxonName
WHEN TaxonName2.RankID = 170 THEN TaxonName2.TaxonName
WHEN TaxonName3.RankID = 170 THEN TaxonName3.TaxonName
WHEN TaxonName4.RankID = 170 THEN TaxonName4.TaxonName
WHEN TaxonName5.RankID = 170 THEN TaxonName5.TaxonName
WHEN TaxonName6.RankID = 170 THEN TaxonName6.TaxonName
END
,
--SubTribuAuthor
CASE WHEN dbo.TaxonName.RankID = 170 THEN dbo.TaxonName.Author
WHEN TaxonName1.RankID = 170 THEN TaxonName1.Author
WHEN TaxonName2.RankID = 170 THEN TaxonName2.Author
WHEN TaxonName3.RankID = 170 THEN TaxonName3.Author
WHEN TaxonName4.RankID = 170 THEN TaxonName4.Author
WHEN TaxonName5.RankID = 170 THEN TaxonName5.Author
WHEN TaxonName6.RankID = 170 THEN TaxonName6.Author
 
END
,
--Genus
CASE WHEN dbo.TaxonName.RankID = 180 THEN dbo.TaxonName.TaxonName
WHEN TaxonName1.RankID = 180 THEN TaxonName1.TaxonName
WHEN TaxonName2.RankID = 180 THEN TaxonName2.TaxonName
WHEN TaxonName3.RankID = 180 THEN TaxonName3.TaxonName
WHEN TaxonName4.RankID = 180 THEN TaxonName4.TaxonName
WHEN TaxonName5.RankID = 180 THEN TaxonName5.TaxonName
END
,
--GenusAuthor
CASE WHEN dbo.TaxonName.RankID = 180 THEN dbo.TaxonName.Author
WHEN TaxonName1.RankID = 180 THEN TaxonName1.Author
WHEN TaxonName2.RankID = 180 THEN TaxonName2.Author
WHEN TaxonName3.RankID = 180 THEN TaxonName3.Author
WHEN TaxonName4.RankID = 180 THEN TaxonName4.Author
WHEN TaxonName5.RankID = 180 THEN TaxonName5.Author
END
,
--Subgenus
CASE WHEN dbo.TaxonName.RankID = 190 THEN dbo.TaxonName.TaxonName
WHEN TaxonName1.RankID = 190 THEN TaxonName1.TaxonName
WHEN TaxonName2.RankID = 190 THEN TaxonName2.TaxonName
WHEN TaxonName3.RankID = 190 THEN TaxonName3.TaxonName
WHEN TaxonName4.RankID = 190 THEN TaxonName4.TaxonName
END
,
--SubgenusAuthor
CASE WHEN dbo.TaxonName.RankID = 190 THEN dbo.TaxonName.Author
WHEN TaxonName1.RankID = 190 THEN TaxonName1.Author
WHEN TaxonName2.RankID = 190 THEN TaxonName2.Author
WHEN TaxonName3.RankID = 190 THEN TaxonName3.Author
WHEN TaxonName4.RankID = 190 THEN TaxonName4.Author
END
,	
--Section
CASE WHEN dbo.TaxonName.RankID = 200 THEN dbo.TaxonName.TaxonName
WHEN TaxonName1.RankID = 200 THEN TaxonName1.TaxonName
WHEN TaxonName2.RankID = 200 THEN TaxonName2.TaxonName
WHEN TaxonName3.RankID = 200 THEN TaxonName3.TaxonName
END
,
--SectionAuthor
CASE WHEN dbo.TaxonName.RankID = 200 THEN dbo.TaxonName.Author
WHEN TaxonName1.RankID = 200 THEN TaxonName1.Author
WHEN TaxonName2.RankID = 200 THEN TaxonName2.Author
WHEN TaxonName3.RankID = 200 THEN TaxonName3.Author
END
,
--Species
CASE WHEN dbo.TaxonName.RankID = 220 THEN dbo.TaxonName.TaxonName
WHEN TaxonName1.RankID = 220 THEN TaxonName1.TaxonName
WHEN TaxonName2.RankID = 220 THEN TaxonName2.TaxonName
END
,
--SpeciesAuthor
CASE WHEN dbo.TaxonName.RankID = 220 THEN dbo.TaxonName.Author
WHEN TaxonName1.RankID = 220 THEN TaxonName1.Author
WHEN TaxonName2.RankID = 220 THEN TaxonName2.Author
END
,
--Subspecies
CASE WHEN dbo.TaxonName.RankID = 230 THEN dbo.TaxonName.TaxonName
WHEN taxonName1.RankID = 230 THEN dbo.TaxonName.TaxonName
END
,
--Subspecies Author
CASE WHEN dbo.TaxonName.RankID = 230 THEN dbo.TaxonName.Author
WHEN TaxonName1.RankID = 230 THEN dbo.TaxonName.Author
END
,
--Variety
CASE WHEN dbo.TaxonName.RankID = 240 THEN dbo.TaxonName.TaxonName
END
,
--Variety Author
CASE WHEN dbo.TaxonName.RankID = 240 THEN dbo.TaxonName.Author
END
,
dbo.Agent.FirstName,
dbo.Agent.LastName,
dbo.Determination.[date],
dbo.Determination.TypeStatusName,
dbo.Determination.DeterminationID
FROM
(
(
dbo.TaxonName
INNER JOIN dbo.Determination
ON dbo.TaxonName.TaxonNameID = dbo.Determination.TaxonNameID
)
INNER JOIN dbo.CollectionObject
ON dbo.Determination.BiologicalObjectID = dbo.CollectionObject.CollectionObjectID
) LEFT JOIN dbo.Agent
ON dbo.Determination.DeterminerID = dbo.Agent.AgentID
LEFT OUTER JOIN dbo.TaxonName as TaxonName1 on dbo.TaxonName.parentTaxonNameID = TaxonName1.TaxonNameID
LEFT OUTER JOIN dbo.TaxonName as TaxonName2 on TaxonName1.parentTaxonNameID = TaxonName2.TaxonNameID
LEFT OUTER JOIN dbo.TaxonName as TaxonName3 on TaxonName2.parentTaxonNameID = TaxonName3.TaxonNameID
LEFT OUTER JOIN dbo.TaxonName as TaxonName4 on TaxonName3.parentTaxonNameID = TaxonName4.TaxonNameID
LEFT OUTER JOIN dbo.TaxonName as TaxonName5 on TaxonName4.parentTaxonNameID = TaxonName5.TaxonNameID
LEFT OUTER JOIN dbo.TaxonName as TaxonName6 on TaxonName5.parentTaxonNameID = TaxonName6.TaxonNameID
LEFT OUTER JOIN dbo.TaxonName as TaxonName7 on TaxonName6.parentTaxonNameID = TaxonName7.TaxonNameID
LEFT OUTER JOIN dbo.TaxonName as TaxonName8 on TaxonName7.parentTaxonNameID = TaxonName8.TaxonNameID
LEFT OUTER JOIN dbo.TaxonName as TaxonName9 on TaxonName8.parentTaxonNameID = TaxonName9.TaxonNameID
LEFT OUTER JOIN dbo.TaxonName as TaxonName10 on TaxonName9.parentTaxonNameID = TaxonName10.TaxonNameID
WHERE dbo.determination.[current] = 0
INSERT INTO @OtrasDet
(
CollectionObjectID,
Kingdom1,
KingdomAuthor1,
Class1,
ClassAuthor1,
[Order1],
OrderAuthor1,
Family1,
FamilyAuthor1,
SubFamily1,
SubFamilyAuthor1,
Subtribu1,
SubtribuAuthor1,
Genus1,
GenusAuthor1,
Subgenus1,
SubgenusAuthor1,
Section1,
SectionAuthor1,
Species1,
SpeciesAuthor1,
Subspecies1,
SubspeciesAuthor1,
Variety1,
VarietyAuthor1,
DeterminerFirstName1,
DeterminerLastName1,
[DeterminationDate1],
TypeStatusName1,
DeterminationID
)
SELECT
 
CollectionObject_ID,
-- Field_Number,
-- Andes_T,
-- Taxon_ID,
Kingdom_Name,
Kingdom_Author,
Class_Name,
Class_Author,
Order_Name,
Order_Author,
Family_Name,
Family_Author,
SubFamily_Name,
SubFamily_Author,
Subtribu_Name,
Subtribu_Author,
Genus_Name,
Genus_Author,
Subgenus_Name,
Subgenus_Author,
Section_Name,
Section_Author,
Species_Name,
Species_Author,
Subspecies_Name,
Subspecies_Author,
Variety_Name,
Variety_Author,
Determiner_FirstName,
Determiner_LastName,
Determination_Date,
TypeStatus_Name,
Determination_ID
FROM @Taxonomy
RETURN
END