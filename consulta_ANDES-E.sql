USE [ANDES_E]
GO
/****** Object:  UserDefinedFunction [dbo].[ConsultaCompleta]    Script Date: 08/13/2013 10:44:52 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:  	Diego Perico Manrique
-- Create date: 20/Jan/ 2013
/** Description:	Esta función genera una consulta que permite recuperar todos los datos
			almacenados en la base de datos de Specify 5 de la coleccion de Bacterias del
			Museo de Historia Natural Andes Andes-B. Incluye los datos del ejemplar, determinaciones,
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
	--Catalog
	CatalogNumber	VARCHAR (20),
	FieldNumber	VARCHAR (20),
	AndesTnumber1	VARCHAR (20),
	CIMPATtext2	NVARCHAR (300),
	CollectionRemarksRemarks	TEXT,
	CatalogerFirstName	VARCHAR(50),
	CatalogerLastName	VARCHAR (50),
	CataloguedDateSPY	VARCHAR (50),
	CataloguedDate	DATETIME NULL,
	AltCatNumber	VARCHAR (20),
	--Preparation
	--1
	PreparationMethod1	VARCHAR (50),
	FixationMethod1	VARCHAR (50),
	[Count1]	VARCHAR (10),
	PreparedFirstName1	VARCHAR(50),
	PreparedLastName1	VARCHAR (50),
	PreparedDateSPY1	VARCHAR (20),
	PreparedDate1	DATETIME NULL,
	--2
	PreparationMethod2	VARCHAR (50),
	FixationMethod2	VARCHAR (50),
	[Count2]	VARCHAR (10),
	PreparedFirstName2	VARCHAR(50),
	PreparedLastName2	VARCHAR (50),
	PreparedDateSPY2	VARCHAR (20),
	PreparedDate2	DATETIME NULL,
	--3
	PreparationMethod3	VARCHAR (50),
	FixationMethod3	VARCHAR (50),
	[Count3]	VARCHAR (10),
	PreparedFirstName3	VARCHAR(50),
	PreparedLastName3	VARCHAR (50),
	PreparedDateSPY3	VARCHAR (20),
	PreparedDate3	DATETIME NULL,
	--4
	PreparationMethod4	VARCHAR (50),
	FixationMethod4	VARCHAR (50),
	[Count4]	VARCHAR (10),
	PreparedFirstName4	VARCHAR(50),
	PreparedLastName4	VARCHAR (50),
	PreparedDateSPY4	VARCHAR (20),
	PreparedDate4	DATETIME NULL,
	--Determination
	TaxonID	INT,
	--1
	Kingdom1	VARCHAR (10),
	KingdomAuthor1	VARCHAR (50),
	--2
	Phylum1	VARCHAR (10),
	PhylumAuthor1	VARCHAR (50),
	--3
	Class1	VARCHAR (50),
	ClassAuthor1	VARCHAR (50),
	--4
	Subclass1	VARCHAR (50),
	SubclassAuthor1	VARCHAR (50),
	--5
	Infraclass1	VARCHAR (50),
	InfraclassAuthor1	VARCHAR (50),
	--6
	[Order1]	VARCHAR (50),
	OrderAuthor1	VARCHAR (50),
	--7
	Suborder1	VARCHAR (50),
	SuborderAuthor1	VARCHAR (50),
	--8
	Infraorder1	VARCHAR (50),
	InfraorderAuthor	VARCHAR (50),
	--9
	Superfamily	VARCHAR (50),
	SuperfamilyAuthor	VARCHAR (50),
	--10
	Family1	VARCHAR (50),
	FamilyAuthor1	VARCHAR (50),
	--11
	Subfamily1	VARCHAR (50),
	SubfamilyAuthor1	VARCHAR (50),
	--12
	Tribe1	VARCHAR (50),
	TribeAuthor1	VARCHAR (50),
	--13
	Genus1	VARCHAR (50),
	GenusAuthor1	VARCHAR (50),
	--14
	Subgenus1	VARCHAR (50),
	SubgenusAuthor1	VARCHAR (50),
	--15
	Species1	VARCHAR(50),
	SpeciesAuthor1	VARCHAR (100),
	--16
	Subspecies1	VARCHAR (50),
	SubspeciesAuthor1	VARCHAR (100),
	DeterminerFirstName1	VARCHAR (100),
	DeterminerLastName1	VARCHAR (100),
	DeterminationDateSPY1	VARCHAR (10),
	DeterminationDate1	DATETIME NULL,
	TypeStatusName1	VARCHAR (100),
	Currente1 BIT,
	--Segunda Determinación
	--1
	Kingdom2	VARCHAR (10),
	KingdomAuthor2	VARCHAR (50),
	--2
	Phylum2	VARCHAR (10),
	PhylumAuthor2	VARCHAR (50),
	--3
	Class2	VARCHAR (50),
	ClassAuthor2	VARCHAR (50),
	--4
	Subclass2	VARCHAR (50),
	SubclassAuthor2	VARCHAR (50),
	--5
	Infraclass2	VARCHAR (50),
	InfraclassAuthor2	VARCHAR (50),
	--6
	[Order2]	VARCHAR (50),
	OrderAuthor2	VARCHAR (50),
	--7
	Suborder2	VARCHAR (50),
	SuborderAuthor2	VARCHAR (50),
	--8
	Infraorder2	VARCHAR (50),
	InfraorderAuthor2	VARCHAR (50),
	--9
	Superfamily2	VARCHAR (50),
	SuperfamilyAuthor2	VARCHAR (50),
	--10
	Family2	VARCHAR (50),
	FamilyAuthor2	VARCHAR (50),
	--11
	Subfamily2	VARCHAR (50),
	SubfamilyAuthor2	VARCHAR (50),
	--12
	Tribe2	VARCHAR (50),
	TribeAuthor2	VARCHAR (50),
	--13
	Genus2	VARCHAR (50),
	GenusAuthor2	VARCHAR (50),
	--14
	Subgenus2	VARCHAR (50),
	SubgenusAuthor2	VARCHAR (50),
	--15
	Species2	VARCHAR(50),
	SpeciesAuthor2	VARCHAR (100),
	--16
	Subspecies2	VARCHAR (50),
	SubspeciesAuthor2	VARCHAR (100),
	DeterminerFirstName2	VARCHAR (100),
	DeterminerLastName2	VARCHAR (100),
	DeterminationDateSPY2	VARCHAR (10),
	DeterminationDate2	DATETIME NULL,
	TypeStatusName2	VARCHAR (100),
	Currente2	BIT,
	--Tercera Determinación
	--1
	Kingdom3	VARCHAR (10),
	KingdomAuthor3	VARCHAR (50),
	--2
	Phylum3	VARCHAR (10),
	PhylumAuthor3	VARCHAR (50),
	--3
	Class3	VARCHAR (50),
	ClassAuthor3	VARCHAR (50),
	--4
	Subclass3	VARCHAR (50),
	SubclassAuthor3	VARCHAR (50),
	--5
	Infraclass3	VARCHAR (50),
	InfraclassAuthor3	VARCHAR (50),
	--6
	[Order3]	VARCHAR (50),
	OrderAuthor3	VARCHAR (50),
	--7
	Suborder3	VARCHAR (50),
	SuborderAuthor3	VARCHAR (50),
	--8
	Infraorder3	VARCHAR (50),
	InfraorderAuthor3	VARCHAR (50),
	--9
	Superfamily3	VARCHAR (50),
	SuperfamilyAuthor3	VARCHAR (50),
	--10
	Family3	VARCHAR (50),
	FamilyAuthor3	VARCHAR (50),
	--11
	Subfamily3	VARCHAR (50),
	SubfamilyAuthor3	VARCHAR (50),
	--12
	Tribe3	VARCHAR (50),
	TribeAuthor3	VARCHAR (50),
	--13
	Genus3	VARCHAR (50),
	GenusAuthor3	VARCHAR (50),
	--14
	Subgenus3	VARCHAR (50),
	SubgenusAuthor3	VARCHAR (50),
	--15
	Species3	VARCHAR(50),
	SpeciesAuthor3	VARCHAR (100),
	--16
	Subspecies3	VARCHAR (50),
	SubspeciesAuthor3	VARCHAR (100),
	DeterminerFirstName3	VARCHAR (100),
	DeterminerLastName3	VARCHAR (100),
	DeterminationDateSPY3	VARCHAR (10),
	DeterminationDate3	DATETIME NULL,
	TypeStatusName3	VARCHAR (100),
	Currente3	BIT,
	--Collecting
	Continent	VARCHAR (100),
	Country	VARCHAR (100),
	[State]	VARCHAR (100),
	County	VARCHAR (100),
	Locality	TEXT,
	Named	TEXT,
	MinElevation	VARCHAR (10),
	MaxElevation	VARCHAR (10),
	ElevationUnits	VARCHAR (10),
	ElevationAcuracy	VARCHAR (10),
	Latitude1	DECIMAL (12,10),
	Longitude1	DECIMAL (12,10),
	Latitude2	DECIMAL (12,10),
	Longitude2	DECIMAL (12,10),
	[Latitude/LongitudeAccuracy]	VARCHAR (10),
	[Latitude/LongitudeType]	VARCHAR (10),
	LatLongMethod	VARCHAR (20),
	ElevationMethod	VARCHAR (20),
	StartDateSPY	VARCHAR (10),
	StartDate	DATETIME NULL,
	EndDateSPY	VARCHAR (10),
	EndDate	DATETIME NULL,
	StartTime	SMALLINT,
	EndTime	SMALLINT,
	Remarks	TEXT,
	AttractiveREMmethod	VARCHAR (50),
	CollectionRemarksVerbaLocal	TEXT,
	--collector
	CollectorFirstName1	VARCHAR (100),
	CollectorLastName1	VARCHAR (100),
	CollectorMiddle1	VARCHAR (20),
	CollectorFirstName2	VARCHAR (100),
	CollectorLastName2	VARCHAR (100),
	CollectorMiddle2	VARCHAR (20),
	CollectorFirstName3	VARCHAR (100),
	CollectorLastName3	VARCHAR (100),
	CollectorMiddle3	VARCHAR (20),
	CollectorFirstName4	VARCHAR (100),
	CollectorLastName4	VARCHAR (100),
	CollectorMiddle4	VARCHAR (20),
	--attributes
	Sex	VARCHAR (50),
	Stage	VARCHAR (50),
	MorphotypeRemarksText1	NVARCHAR (300)
	

	)
AS BEGIN
	
	DECLARE @taxonomy TABLE
	(
	 CollectionObject_ID	INT,
	 Field_Number	VARCHAR (20),
	 Andes_Tnumber1	VARCHAR (20),
	 CIMPAT_text2	NVARCHAR (300),
	 RemarksCollObj	TEXT,
	 Taxon_ID	INT,
	 --1
	 Kingdom_Name	VARCHAR (10),
	 Kingdom_Author	VARCHAR (50),
	 --2
	 Phylum_name	VARCHAR (50),
	 Phylum_Author	VARCHAR (50),
	 --3
	 Class_Name	VARCHAR (50),
	 Class_Author	VARCHAR (50),
	 --4
	 Subclass_Name	VARCHAR (50),
	 Subclass_Author	VARCHAR (50),
	 --5
	 Infraclass_Name	VARCHAR (50),
	 Infraclass_Author	VARCHAR (50),
	 --6
	 Order_name	VARCHAR (50),
	 Order_Author	VARCHAR (50),
	 --7
	 Suborder_Name	VARCHAR (50),
	 Suborder_Author	VARCHAR (50),
	 --8
	 Infraorder_Name	VARCHAR (50),
	 Infraorder_Author	VARCHAR (50),
	 --9
	 Superfamily_name	VARCHAR (50),
	 Superfamily_Author	VARCHAR (50),
	 --10
	 Family_name	VARCHAR (50),
	 Family_Author	VARCHAR (50),
	 --11
	 SubFamily_name	VARCHAR (50),
	 SubFamily_Author	VARCHAR (50),
	 --12
	 Tribe_Name	VARCHAR (50),
	 Tribe_Author	VARCHAR (50),
	 --13
	 Genus_Name	VARCHAR(50),
	 Genus_Author	VARCHAR (50),
	 --14
	 Subgenus_Name	VARCHAR(50),
	 Subgenus_Author	VARCHAR (50),
	 --15
	 Species_Name	VARCHAR(50),
	 Species_Author	VARCHAR (100),
	 --16
	 Subspecies_Name	VARCHAR (50),
	 Subspecies_Author	VARCHAR (100),
	 Determiner_FirstName	VARCHAR (100),
	 Determiner_LastName	VARCHAR (100),
	 [Determination_Date]	VARCHAR (10),
	 TypeStatus_Name	VARCHAR (100),
	 isCurrent	BIT

	)
	INSERT INTO @taxonomy
	(
	 CollectionObject_ID,
	 Field_Number,
	 Andes_Tnumber1,
	 CIMPAT_text2,
	 RemarksCollObj,
	 Taxon_ID,
	 --1
	 Kingdom_Name,
	 Kingdom_Author,
	 --2
	 Phylum_name,
	 Phylum_Author,
	 --3
	 Class_Name,
	 Class_Author,
	 --4
	 Subclass_Name,
	 Subclass_Author,
	 --5
	 Infraclass_Name,
	 Infraclass_Author,
	 --6
	 Order_name,
	 Order_Author,
	 --7
	 Suborder_Name,
	 Suborder_Author,
	 --8
	 Infraorder_Name,
	 Infraorder_Author,
	 --9
	 Superfamily_name,
	 Superfamily_Author,
	 --10
	 Family_name,
	 Family_Author,
	 --11
	 SubFamily_name,
	 SubFamily_Author,
	 --12
	 Tribe_Name,
	 Tribe_Author,
	 --13
	 Genus_Name,
	 Genus_Author,
	 --14
	 Subgenus_Name,
	 Subgenus_Author,
	 --15
	 Species_Name,
	 Species_Author,
	 --16
	 Subspecies_Name,
	 Subspecies_Author,
	 Determiner_FirstName,
	 Determiner_LastName,
	 [Determination_Date],
	 TypeStatus_Name,
	 isCurrent

	)
	SELECT 
	
	dbo.CollectionObject.CollectionObjectID,
	dbo.CollectionObject.FieldNumber,
	dbo.CollectionObject.Number1,
	dbo.CollectionObject.Text2,
	dbo.CollectionObject.Remarks,
	dbo.TaxonName.TaxonNameID,
	--1
	--Kingdom
	'Animalia',
	--Kingdom Author
	'Linnaeus, 1758',
	--2
	--Phylum
	'Artropoda',
	--Phylum Author
	'Latreille, 1829',
	--3
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
	--4
	--Subclass
	CASE WHEN dbo.TaxonName.RankID = 70 THEN dbo.TaxonName.TaxonName
		WHEN TaxonName1.RankID = 70 THEN TaxonName1.TaxonName
		WHEN TaxonName2.RankID = 70 THEN TaxonName2.TaxonName
		WHEN TaxonName3.RankID = 70 THEN TaxonName3.TaxonName
		WHEN TaxonName4.RankID = 70 THEN TaxonName4.TaxonName
		WHEN TaxonName5.RankID = 70 THEN TaxonName5.TaxonName
		WHEN TaxonName6.RankID = 70 THEN TaxonName6.TaxonName
		WHEN TaxonName7.RankID = 70 THEN TaxonName7.TaxonName
		WHEN TaxonName8.RankID = 70 THEN TaxonName5.TaxonName
		WHEN TaxonName9.RankID = 70 THEN TaxonName6.TaxonName
		WHEN TaxonName10.RankID = 70 THEN TaxonName7.TaxonName
	END
	,	
	--SubclassAuthor
	CASE WHEN dbo.TaxonName.RankID = 70 THEN dbo.TaxonName.Author
		WHEN TaxonName1.RankID = 70 THEN TaxonName1.Author
		WHEN TaxonName2.RankID = 70 THEN TaxonName2.Author
		WHEN TaxonName3.RankID = 70 THEN TaxonName3.Author
		WHEN TaxonName4.RankID = 70 THEN TaxonName4.Author
		WHEN TaxonName5.RankID = 70 THEN TaxonName5.Author
		WHEN TaxonName6.RankID = 70 THEN TaxonName6.Author
		WHEN TaxonName7.RankID = 70 THEN TaxonName7.Author
		WHEN TaxonName8.RankID = 70 THEN TaxonName8.Author
		WHEN TaxonName9.RankID = 70 THEN TaxonName9.Author
		WHEN TaxonName10.RankID = 70 THEN TaxonName10.Author
	END
	,
	--5
	--Infraclass
	CASE WHEN dbo.TaxonName.RankID = 80 THEN dbo.TaxonName.TaxonName
		WHEN TaxonName1.RankID = 80 THEN TaxonName1.TaxonName
		WHEN TaxonName2.RankID = 80 THEN TaxonName2.TaxonName
		WHEN TaxonName3.RankID = 80 THEN TaxonName3.TaxonName
		WHEN TaxonName4.RankID = 80 THEN TaxonName4.TaxonName
		WHEN TaxonName5.RankID = 80 THEN TaxonName5.TaxonName
		WHEN TaxonName6.RankID = 80 THEN TaxonName6.TaxonName
		WHEN TaxonName7.RankID = 80 THEN TaxonName7.TaxonName
		WHEN TaxonName8.RankID = 80 THEN TaxonName5.TaxonName
		WHEN TaxonName9.RankID = 80 THEN TaxonName6.TaxonName
	END
	,	
	--InfraclassAuthor
	CASE WHEN dbo.TaxonName.RankID = 80 THEN dbo.TaxonName.Author
		WHEN TaxonName1.RankID = 80 THEN TaxonName1.Author
		WHEN TaxonName2.RankID = 80 THEN TaxonName2.Author
		WHEN TaxonName3.RankID = 80 THEN TaxonName3.Author
		WHEN TaxonName4.RankID = 80 THEN TaxonName4.Author
		WHEN TaxonName5.RankID = 80 THEN TaxonName5.Author
		WHEN TaxonName6.RankID = 80 THEN TaxonName6.Author
		WHEN TaxonName7.RankID = 80 THEN TaxonName7.Author
		WHEN TaxonName8.RankID = 80 THEN TaxonName8.Author
		WHEN TaxonName9.RankID = 80 THEN TaxonName9.Author
	END
	,
	--6	
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
	--7
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
	,
	--8
	--Infraorder
	CASE WHEN dbo.TaxonName.RankID = 120 THEN dbo.TaxonName.TaxonName
		WHEN TaxonName1.RankID = 120 THEN TaxonName1.TaxonName
		WHEN TaxonName2.RankID = 120 THEN TaxonName2.TaxonName
		WHEN TaxonName3.RankID = 120 THEN TaxonName3.TaxonName
		WHEN TaxonName4.RankID = 120 THEN TaxonName4.TaxonName
		WHEN TaxonName5.RankID = 120 THEN TaxonName5.TaxonName
		WHEN TaxonName6.RankID = 120 THEN TaxonName6.TaxonName
		WHEN TaxonName7.RankID = 120 THEN TaxonName7.TaxonName
		WHEN TaxonName8.RankID = 120 THEN TaxonName8.TaxonName
		WHEN TaxonName9.RankID = 120 THEN TaxonName9.TaxonName
	END
	,
	--Infraorder Author
	CASE WHEN dbo.TaxonName.RankID = 120 THEN dbo.TaxonName.Author
		WHEN TaxonName1.RankID = 120 THEN TaxonName1.Author
		WHEN TaxonName2.RankID = 120 THEN TaxonName2.Author
		WHEN TaxonName3.RankID = 120 THEN TaxonName3.Author
		WHEN TaxonName4.RankID = 120 THEN TaxonName4.Author
		WHEN TaxonName5.RankID = 120 THEN TaxonName5.Author
		WHEN TaxonName6.RankID = 120 THEN TaxonName6.Author
		WHEN TaxonName7.RankID = 120 THEN TaxonName7.Author
		WHEN TaxonName8.RankID = 120 THEN TaxonName8.Author
		WHEN TaxonName9.RankID = 120 THEN TaxonName9.Author
	END
	,
	--9
	--Superfamily
	CASE WHEN dbo.TaxonName.RankID = 130 THEN dbo.TaxonName.TaxonName
		WHEN TaxonName1.RankID = 130 THEN TaxonName1.TaxonName
		WHEN TaxonName2.RankID = 130 THEN TaxonName2.TaxonName
		WHEN TaxonName3.RankID = 130 THEN TaxonName3.TaxonName
		WHEN TaxonName4.RankID = 130 THEN TaxonName4.TaxonName
		WHEN TaxonName5.RankID = 130 THEN TaxonName5.TaxonName
		WHEN TaxonName6.RankID = 130 THEN TaxonName6.TaxonName
		WHEN TaxonName7.RankID = 130 THEN TaxonName7.TaxonName
		WHEN TaxonName8.RankID = 130 THEN TaxonName8.TaxonName
	END
	,
	--FamilyAuthor
	CASE WHEN dbo.TaxonName.RankID = 130 THEN dbo.TaxonName.Author
		WHEN TaxonName1.RankID = 130 THEN TaxonName1.Author
		WHEN TaxonName2.RankID = 130 THEN TaxonName2.Author
		WHEN TaxonName3.RankID = 130 THEN TaxonName3.Author
		WHEN TaxonName4.RankID = 130 THEN TaxonName4.Author
		WHEN TaxonName5.RankID = 130 THEN TaxonName5.Author
		WHEN TaxonName6.RankID = 130 THEN TaxonName6.Author
		WHEN TaxonName7.RankID = 130 THEN TaxonName7.Author
		WHEN TaxonName8.RankID = 130 THEN TaxonName8.Author
	END
	,
	--10
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
	--11
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
	--12
	--Tribe
	CASE WHEN dbo.TaxonName.RankID = 160 THEN dbo.TaxonName.TaxonName
		WHEN TaxonName1.RankID = 160 THEN TaxonName1.TaxonName
		WHEN TaxonName2.RankID = 160 THEN TaxonName2.TaxonName
		WHEN TaxonName3.RankID = 160 THEN TaxonName3.TaxonName
		WHEN TaxonName4.RankID = 160 THEN TaxonName4.TaxonName
		WHEN TaxonName5.RankID = 160 THEN TaxonName5.TaxonName
		WHEN TaxonName6.RankID = 160 THEN TaxonName6.TaxonName
	END
	,
	--TribeAuthor
		CASE WHEN dbo.TaxonName.RankID = 160 THEN dbo.TaxonName.Author
		WHEN TaxonName1.RankID = 160 THEN TaxonName1.Author
		WHEN TaxonName2.RankID = 160 THEN TaxonName2.Author
		WHEN TaxonName3.RankID = 160 THEN TaxonName3.Author
		WHEN TaxonName4.RankID = 160 THEN TaxonName4.Author
		WHEN TaxonName5.RankID = 160 THEN TaxonName5.Author
		WHEN TaxonName6.RankID = 160 THEN TaxonName6.Author

	END
	,
	--13
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
	--14
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
	--15
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
	--16
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
	dbo.Agent.FirstName,
	dbo.Agent.LastName,
	dbo.Determination.[date],
	dbo.Determination.TypeStatusName,
	dbo.Determination.[Current]
	
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
	Min_Elevation	VARCHAR (10),
	Max_Elevation	VARCHAR (10),
	Elevation_Units	VARCHAR (10),
	Elevation_Accuracy	VARCHAR (10),
	Latitude_1	DECIMAL (12,10),
	Longitude_1	DECIMAL (12,10),
	Latitude_2	DECIMAL (12,10),
	Longitude_2	DECIMAL (12, 10),
	[Latitude/Longitude_Accuracy]	VARCHAR (10),
	[Latitude/Longitude_Type]	VARCHAR (10),
	LatLong_Method	VARCHAR (20),
	Elevation_Method	VARCHAR (20),
	Star_Date	VARCHAR (10),
	End_Date	VARCHAR (10),
	Start_Time	SMALLINT,
	End_Time	SMALLINT,
	CollectionRemarks_VerbatimLoca	TEXT,
	AttractiveREM_method	VARCHAR (50),
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
	Collector1Middle	VARCHAR (20)
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
WHERE (((dbo.Collectors.[Order])=1))

DECLARE @collector2 TABLE
	(
	CollectionObjectC2_ID	INT,
	Collector2FirstName	VARCHAR (100),
	Collector2LastName	VARCHAR (100),
	Collector2Middle	VARCHAR (20)
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
WHERE (((dbo.Collectors.[Order])=2))

DECLARE @collector3 TABLE
	(
	CollectionObjectC3_ID	INT,
	Collector3FirstName	VARCHAR (100),
	Collector3LastName	VARCHAR (100),
	Collector3Middle	VARCHAR (20)
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
WHERE (((dbo.Collectors.[Order])=3))

DECLARE @collector4 TABLE
	(
	CollectionObjectC4_ID	INT,
	Collector4FirstName	VARCHAR (100),
	Collector4LastName	VARCHAR (100),
	Collector4Middle	VARCHAR (20)
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
WHERE (((dbo.Collectors.[Order])=4))

	
DECLARE @collectors TABLE
	(
	CollectionObject_ID2	INT,
	CollectorFirstName_1	VARCHAR (100),
	CollectorLastName_1	VARCHAR (100),
	CollectorMiddle_1	VARCHAR (20),
	CollectorFirstName_2	VARCHAR (100),
	CollectorLastName_2	VARCHAR (100),
	CollectorMiddle_2	VARCHAR (20),
	CollectorFirstName_3	VARCHAR (100),
	CollectorLastName_3	VARCHAR (100),
	CollectorMiddle_3	VARCHAR (20),
	CollectorFirstName_4	VARCHAR (100),
	CollectorLastName_4	VARCHAR (100),
	CollectorMiddle_4	VARCHAR (20)
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
Sexatt	VARCHAR (50),
Stageatt	VARCHAR (50),
MorphotypeRemarks_Text1	NVARCHAR (300)
	)

INSERT INTO @Attributes
	(
CollectionObject_ID3,
sexatt,
Stageatt,
MorphotypeRemarks_Text1
)
SELECT

dbo.BiologicalObjectAttributes.BiologicalObjectAttributesID,
dbo.BiologicalObjectAttributes.sex,
dbo.BiologicalObjectAttributes.Stage,
dbo.BiologicalObjectAttributes.Text1

from dbo.BiologicalObjectAttributes

DECLARE @PrimPrep TABLE
	(
	PrimPrepMethodID	INT,
	CatalogNumberPrep	VARCHAR (50)
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
	Catalog_Number	VARCHAR (20),
	Preparation_Method	VARCHAR (50),
	Fixation_Method	VARCHAR (50),
	CatalogerFirst_Name	VARCHAR(50),
	CatalogerLast_Name	VARCHAR (50),
	Catalogued_Date	VARCHAR (50),
	Count_1	VARCHAR (10),
	PreviousNumber VARCHAR (20)
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
CASE WHEN  CollectionObject_1.PreparationMethod IS NULL THEN 'Pinned'
	ELSE CollectionObject_1.PreparationMethod
	END
	as PreparationMethod
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
	Prepared_Date	VARCHAR (20),
	PreparedBy_FirstName	VARCHAR (50),
	PreparedBy_LastName	VARCHAR (50)
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
	
INSERT INTO @Completa
	(
	CollectionObjectID,
	--Catalog
	CatalogNumber,
	FieldNumber,
	AndesTnumber1,
	CIMPATtext2,
	CollectionRemarksRemarks,
	CatalogerFirstName,
	CatalogerLastName,
	CataloguedDateSPY,
	--CataloguedDate,
	AltCatNumber,
	--Preparation
	--1
	PreparationMethod1,
	FixationMethod1,
	[Count1],
	PreparedFirstName1,
	PreparedLastName1,
	PreparedDateSPY1,
	--PreparedDate1,
	--2
	PreparationMethod2,
	FixationMethod2,
	[Count2],
	PreparedFirstName2,
	PreparedLastName2,
	PreparedDateSPY2,
	--PreparedDate2,
	--3
	PreparationMethod3,
	FixationMethod3,
	[Count3],
	PreparedFirstName3,
	PreparedLastName3,
	PreparedDateSPY3,
	--PreparedDate3,
	--4
	PreparationMethod4,
	FixationMethod4,
	[Count4],
	PreparedFirstName4,
	PreparedLastName4,
	PreparedDateSPY4,
	--PreparedDate4,
	--Determination
	TaxonID,
	--1
	Kingdom1,
	KingdomAuthor1,
	--2
	Phylum1,
	PhylumAuthor1,
	--3
	Class1,
	ClassAuthor1,
	--4
	Subclass1,
	SubclassAuthor1,
	--5
	Infraclass1,
	InfraclassAuthor1,
	--6
	[Order1],
	OrderAuthor1,
	--7
	Suborder1,
	SuborderAuthor1,
	--8
	Infraorder1,
	InfraorderAuthor,
	--9
	Superfamily,
	SuperfamilyAuthor,
	--10
	Family1,
	FamilyAuthor1,
	--11
	Subfamily1,
	SubfamilyAuthor1,
	--12
	Tribe1,
	TribeAuthor1,
	--13
	Genus1,
	GenusAuthor1,
	--14
	Subgenus1,
	SubgenusAuthor1,
	--15
	Species1,
	SpeciesAuthor1,
	--16
	Subspecies1,
	SubspeciesAuthor1,
	DeterminerFirstName1,
	DeterminerLastName1,
	DeterminationDateSPY1,
	--DeterminationDate1,
	TypeStatusName1,
	Currente1,
	--Segunda Determinación
	--1
	Kingdom2,
	KingdomAuthor2,
	--2
	Phylum2,
	PhylumAuthor2,
	--3
	Class2,
	ClassAuthor2,
	--4
	Subclass2,
	SubclassAuthor2,
	--5
	Infraclass2,
	InfraclassAuthor2,
	--6
	[Order2],
	OrderAuthor2,
	--7
	Suborder2,
	SuborderAuthor2,
	--8
	Infraorder2,
	InfraorderAuthor2,
	--9
	Superfamily2,
	SuperfamilyAuthor2,
	--10
	Family2,
	FamilyAuthor2,
	--11
	Subfamily2,
	SubfamilyAuthor2,
	--12
	Tribe2,
	TribeAuthor2,
	--13
	Genus2,
	GenusAuthor2,
	--14
	Subgenus2,
	SubgenusAuthor2,
	--15
	Species2,
	SpeciesAuthor2,
	--16
	Subspecies2,
	SubspeciesAuthor2,
	DeterminerFirstName2,
	DeterminerLastName2,
	DeterminationDateSPY2,
	--DeterminationDate2,
	TypeStatusName2,
	Currente2,
	--Tercera Determinación
	--1
	Kingdom3,
	KingdomAuthor3,
	--2
	Phylum3,
	PhylumAuthor3,
	--3
	Class3,
	ClassAuthor3,
	--4
	Subclass3,
	SubclassAuthor3,
	--5
	Infraclass3,
	InfraclassAuthor3,
	--6
	[Order3],
	OrderAuthor3,
	--7
	Suborder3,
	SuborderAuthor3,
	--8
	Infraorder3,
	InfraorderAuthor3,
	--9
	Superfamily3,
	SuperfamilyAuthor3,
	--10
	Family3,
	FamilyAuthor3,
	--11
	Subfamily3,
	SubfamilyAuthor3,
	--12
	Tribe3,
	TribeAuthor3,
	--13
	Genus3,
	GenusAuthor3,
	--14
	Subgenus3,
	SubgenusAuthor3,
	--15
	Species3,
	SpeciesAuthor3,
	--16
	Subspecies3,
	SubspeciesAuthor3,
	DeterminerFirstName3,
	DeterminerLastName3,
	DeterminationDateSPY3,
	--DeterminationDate3,
	TypeStatusName3,
	Currente3,
	--Collecting
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
	--StartDate,
	EndDateSPY,
	--EndDate,
	StartTime,
	EndTime,
	Remarks,
	AttractiveREMmethod,
	CollectionRemarksVerbaLocal,
	--collector
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
	--attributes
	Sex,
	Stage,
	MorphotypeRemarksText1

	)
	
	SELECT

	CollectionObject_ID,
	--catalog
	Catalog_Number,
	Field_Number,
	Andes_Tnumber1,
	CIMPAT_text2,
	RemarksCollObj,
	CatalogerFirst_Name,
	CatalogerLast_Name,
	Catalogued_Date,
	--'',
	PreviousNumber,
	--Preparation
	--1
	Preparation_Method,
	Fixation_Method,
	Count_1,
	PreparedBy_FirstName,
	PreparedBy_LastName,
	Prepared_Date,
	--'',
	--2
	NULL,
	NULL,
	NULL,
	NULL,
	NULL,
	NULL,
	--'',
	--3
	NULL,
	NULL,
	NULL,
	NULL,
	NULL,
	NULL,
	--'',
	--4
	NULL,
	NULL,
	NULL,
	NULL,
	NULL,
	NULL,
	--'',
	--Determination
	Taxon_ID,
	 --1
	 Kingdom_Name,
	 Kingdom_Author,
	 --2
	 Phylum_name,
	 Phylum_Author,
	 --3
	 Class_Name,
	 Class_Author,
	 --4
	 Subclass_Name,
	 Subclass_Author,
	 --5
	 Infraclass_Name,
	 Infraclass_Author,
	 --6
	 Order_name,
	 Order_Author,
	 --7
	 Suborder_Name,
	 Suborder_Author,
	 --8
	 Infraorder_Name,
	 Infraorder_Author,
	 --9
	 Superfamily_name,
	 Superfamily_Author,
	 --10
	 Family_name,
	 Family_Author,
	 --11
	 SubFamily_name,
	 SubFamily_Author,
	 --12
	 Tribe_Name,
	 Tribe_Author,
	 --13
	 Genus_Name,
	 Genus_Author,
	 --14
	 Subgenus_Name,
	 Subgenus_Author,
	 --15
	 Species_Name,
	 Species_Author,
	 --16
	 Subspecies_Name,
	 Subspecies_Author,
	 Determiner_FirstName,
	 Determiner_LastName,
	 [Determination_Date],
	 --'',
	 TypeStatus_Name,
	 isCurrent,
	--Segunda Determinación
	--1
	NULL,
	NULL,
	--2
	NULL,
	NULL,
	--3
	NULL,
	NULL,
	--4
	NULL,
	NULL,
	--5
	NULL,
	NULL,
	--6
	NULL,
	NULL,
	--7
	NULL,
	NULL,
	--8
	NULL,
	NULL,
	--9
	NULL,
	NULL,
	--10
	NULL,
	NULL,
	--11
	NULL,
	NULL,
	--12
	NULL,
	NULL,
	--13
	NULL,
	NULL,
	--14
	NULL,
	NULL,
	--15
	NULL,
	NULL,
	--16
	NULL,
	NULL,
	NULL,
	NULL,
	NULL,
	--'',
	NULL,
	NULL,
	--Tercera Determinación
	--1
	NULL,
	NULL,
	--2
	NULL,
	NULL,
	--3
	NULL,
	NULL,
	--4
	NULL,
	NULL,
	--5
	NULL,
	NULL,
	--6
	NULL,
	NULL,
	--7
	NULL,
	NULL,
	--8
	NULL,
	NULL,
	--9
	NULL,
	NULL,
	--10
	NULL,
	NULL,
	--11
	NULL,
	NULL,
	--12
	NULL,
	NULL,
	--13
	NULL,
	NULL,
	--14
	NULL,
	NULL,
	--15
	NULL,
	NULL,
	--16
	NULL,
	NULL,
	NULL,
	NULL,
	NULL,
	--'',
	NULL,
	NULL,
	--collecting
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
	--'',
	End_Date,
	--'',
	Start_Time,
	End_Time,
	Remarks,
	AttractiveREM_method,
	CollectionRemarks_VerbatimLoca,
	--colectors
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
	--Attributes
	sexatt,
	Stageatt,
	MorphotypeRemarks_Text1

	
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
	
RETURN
END