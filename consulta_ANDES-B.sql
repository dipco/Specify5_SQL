USE [ANDES_B]
GO
/****** Object:  UserDefinedFunction [dbo].[ConsultaCompleta]    Script Date: 08/13/2013 10:31:52 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:  	Diego Perico Manrique
-- Create date: 15/jan/ 2013
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
	PosRevcoNumber1	FLOAT,
	AndesTNumber2	FLOAT,
	AmbienteText1	NVARCHAR (300),
	CepaText2	NVARCHAR (300),
	CarMolecuRemarks	NTEXT,
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
	Subkingdom1	VARCHAR (50),
	SubkingdomAuthor1	VARCHAR (150),
	--3
	Phyllum1	VARCHAR (50),
	PhyllumAuthor1	VARCHAR (150),
	--4
	Class1	VARCHAR (150),
	ClassAuthor1	VARCHAR (150),
	--5
	Subclass1	VARCHAR (150),
	SubClassAuthor1 VARCHAR (150),
	--6
	Order1	VARCHAR (150),
	OrderAuthor1	VARCHAR (150),
	--7
	Suborder1	VARCHAR (150),
	SuborderAuthor1	VARCHAR (150),
	--8
	Family1	VARCHAR (150),
	FamilyAuthor1	VARCHAR (150),
	--9
	Tribe1	VARCHAR (150),
	TribeAuthor1	VARCHAR (150),
	--7
	Genus1	VARCHAR (150),
	GenusAuthor1	VARCHAR (150),
	--8
	Species1	VARCHAR(150),
	SpeciesAuthor1	VARCHAR (100),
	--9
	Variety1	VARCHAR (150),
	VarietyAuthor1	VARCHAR (100),
	
	DeterminerFirstName1	VARCHAR (100),
	DeterminerLastName1	VARCHAR (100),
	DeterminationDateSPY1	INT,
	DeterminationDate1	DATETIME NULL,
	TypeStatusName1	VARCHAR (100),
	Current1	BIT,
	Text1	VARCHAR (150),
	Text2	VARCHAR (150),
	DeterminationRemarks	NTEXT,
	--Segunda Determinación
	
	--1
	Kingdom2	VARCHAR (50),
	KingdomAuthor2	VARCHAR (150),
	--2
	Subkingdom2	VARCHAR (50),
	SubkingdomAuthor2	VARCHAR (150),
	--3
	Phyllum2	VARCHAR (50),
	PnyllumAuthor2	VARCHAR (150),
	--4
	Class2	VARCHAR (150),
	ClassAuthor2	VARCHAR (150),
	--5
	Subclass2	VARCHAR (150),
	SubClassAuthor2 VARCHAR (150),
	--6
	Order2	VARCHAR (150),
	OrderAuthor2	VARCHAR (150),
	--7
	Suborder2	VARCHAR (150),
	SuborderAuthor2 VARCHAR (150),
	--8
	Family2	VARCHAR (150),
	FamilyAuthor2	VARCHAR (150),
	--9
	Tribe2	VARCHAR (150),
	TribeAuthor2	VARCHAR (150),
	--7
	Genus2	VARCHAR (150),
	GenusAuthor2	VARCHAR (150),
	--8
	Species2	VARCHAR(150),
	SpeciesAuthor2	VARCHAR (100),
	--9
	Variety2	VARCHAR (150),
	VarietyAuthor2	VARCHAR (100),
	
	DeterminerFirstName2	VARCHAR (100),
	DeterminerLastName2	VARCHAR (100),
	DeterminationDateSPY2	INT,
	DeterminationDate2	DATETIME NULL,
	TypeStatusName2	VARCHAR (100),
	Current2	BIT,
	--Tercera Determinación
	--1
	Kingdom3	VARCHAR (50),
	KingdomAuthor3	VARCHAR (150),
	--2
	Subkingdom3	VARCHAR (50),
	SubkingdomAuthor3	VARCHAR (150),
	--3
	Phyllum3	VARCHAR (50),
	PnyllumAuthor3	VARCHAR (150),
	--4
	Class3	VARCHAR (150),
	ClassAuthor3	VARCHAR (150),
	--5
	Subclass3	VARCHAR (150),
	SubClassAuthor3 VARCHAR (150),
	--6
	Order3	VARCHAR (150),
	OrderAuthor3	VARCHAR (150),
	--7
	Suborder3	VARCHAR (150),
	SuborderAuthor3 VARCHAR (150),
	--8
	Family3	VARCHAR (150),
	FamilyAuthor3	VARCHAR (150),
	--9
	Tribe3	VARCHAR (150),
	TribeAuthor3	VARCHAR (150),
	--7
	Genus3	VARCHAR (150),
	GenusAuthor3	VARCHAR (150),
	--8
	Species3	VARCHAR(150),
	SpeciesAuthor3	VARCHAR (100),
	--9
	Variety3	VARCHAR (150),
	VarietyAuthor3	VARCHAR (100),
	
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
	CarMolecularRemarks	NTEXT,
	CarMicroText1	VARCHAR (300),
	CarMicroText2	VARCHAR (300),
	OtrosText3	VARCHAR (300)

	)
AS BEGIN
	
	DECLARE @taxonomy TABLE
	(
	 CollectionObject_ID	INT,
	 Field_Number	VARCHAR (100),
	 PosRevcoNumber_1	FLOAT,
	 AndesTNumber_2	FLOAT,
	 AmbienteText_1	NVARCHAR (300),
	 CepaText_2	NVARCHAR (300),
	 RemarksCollObj	TEXT,
	 Taxon_ID	INT,
	 --1
	 Kingdom_Name	VARCHAR (50),
	 Kingdom_Author	VARCHAR (150),
	 --2
	 Subkingdom_Name	VARCHAR (150),
	 Subkingdom_Author	VARCHAR (150),
	 --3
	 Phyllum_name	VARCHAR (150),
	 Phyllum_Author	VARCHAR (150),
	 --4
	 Class_Name	VARCHAR (150),
	 Class_Author	VARCHAR (150),
	 --5
	 Subclass_Name	VARCHAR (150),
	 Subclass_Author	VARCHAR (150),
	 --6
	 Order_Name	VARCHAR (150),
	 Order_Author	VARCHAR (150),
	 --7
	 Suborder_Name	VARCHAR (150),
	 Suborder_Author	VARCHAR (150),
	 --8
	 Family_name	VARCHAR (150),
	 Family_Author	VARCHAR (150),
	 --9
	 Tribe_name	VARCHAR (150),
	 Tribe_Author	VARCHAR (150),
	 --7
	 Genus_Name	VARCHAR(150),
	 Genus_Author	VARCHAR (150),
	 --8
	 Species_Name	VARCHAR(150),
	 Species_Author	VARCHAR (100),
	 --9
	 Variety_Name	VARCHAR (150),
	 Variety_Author	VARCHAR (100),
	 
	 Determiner_FirstName	VARCHAR (100),
	 Determiner_LastName	VARCHAR (100),
	 [Determination_Date]	INT,
	 TypeStatus_Name	VARCHAR (100),
	 isCurrent	BIT,
	 Text_1	VARCHAR (150),
	 Text_2	VARCHAR (150),
	 Determination_Remarks	NTEXT

	)
	INSERT INTO @taxonomy
	(
	 CollectionObject_ID,
	 Field_Number,
	 PosRevcoNumber_1,
	 AndesTNumber_2,
	 AmbienteText_1,
	 CepaText_2,
	 RemarksCollObj,
	 Taxon_ID,
	 --1,
	 Kingdom_Name,
	 Kingdom_Author,
	 --2
	 Subkingdom_Name,
	 Subkingdom_Author,
	 --3
	 Phyllum_name,
	 Phyllum_Author,
	 --4
	 Class_Name,
	 Class_Author,
	 --5
	 Subclass_Name,
	 Subclass_Author,
	 --6
	 Order_name,
	 Order_Author,
	 --7
	 Suborder_Name,
	 Suborder_Author,
	 --8
	 Family_name,
	 Family_Author,
	 --6
	 Tribe_name,
	 Tribe_Author,
	 --7
	 Genus_Name,
	 Genus_Author,
	 --8
	 Species_Name,
	 Species_Author,
	 --9
	 Variety_Name,
	 Variety_Author,

	--Determiner,
	 Determiner_FirstName,
	 Determiner_LastName,
	 [Determination_Date],
	 TypeStatus_Name,
	 isCurrent,
	 Text_1,
	 Text_2,
	 Determination_Remarks

	)
	SELECT 
	
	dbo.CollectionObject.CollectionObjectID,
	dbo.CollectionObject.FieldNumber,
	dbo.CollectionObject.Number1,
	dbo.CollectionObject.Number2,
	dbo.CollectionObject.Text1,
	dbo.CollectionObject.Text2,
	dbo.CollectionObject.Remarks,
	dbo.TaxonName.TaxonNameID,
	
	--1
	--Kingdom
	CASE WHEN dbo.TaxonName.RankID = 10 THEN dbo.TaxonName.TaxonName
		WHEN TaxonName1.RankID = 10 THEN TaxonName1.TaxonName
		WHEN TaxonName2.RankID = 10 THEN TaxonName2.TaxonName
		WHEN TaxonName3.RankID = 10 THEN TaxonName3.TaxonName
		WHEN TaxonName4.RankID = 10 THEN TaxonName4.TaxonName
		WHEN TaxonName5.RankID = 10 THEN TaxonName5.TaxonName
		WHEN TaxonName6.RankID = 10 THEN TaxonName6.TaxonName
		WHEN TaxonName7.RankID = 10 THEN TaxonName7.TaxonName
		WHEN TaxonName8.RankID = 10 THEN TaxonName5.TaxonName
		WHEN TaxonName9.RankID = 10 THEN TaxonName6.TaxonName
	END
	,	
	--KingdomAuthor
	CASE WHEN dbo.TaxonName.RankID = 10 THEN dbo.TaxonName.Author
		WHEN TaxonName1.RankID = 10 THEN TaxonName1.Author
		WHEN TaxonName2.RankID = 10 THEN TaxonName2.Author
		WHEN TaxonName3.RankID = 10 THEN TaxonName3.Author
		WHEN TaxonName4.RankID = 10 THEN TaxonName4.Author
		WHEN TaxonName5.RankID = 10 THEN TaxonName5.Author
		WHEN TaxonName6.RankID = 10 THEN TaxonName6.Author
		WHEN TaxonName7.RankID = 10 THEN TaxonName7.Author
		WHEN TaxonName8.RankID = 10 THEN TaxonName8.Author
		WHEN TaxonName9.RankID = 10 THEN TaxonName9.Author
	END
	,
	--2
	--Submingdom
	CASE WHEN dbo.TaxonName.RankID = 20 THEN dbo.TaxonName.TaxonName
		WHEN TaxonName1.RankID = 20 THEN TaxonName1.TaxonName
		WHEN TaxonName2.RankID = 20 THEN TaxonName2.TaxonName
		WHEN TaxonName3.RankID = 20 THEN TaxonName3.TaxonName
		WHEN TaxonName4.RankID = 20 THEN TaxonName4.TaxonName
		WHEN TaxonName5.RankID = 20 THEN TaxonName5.TaxonName
		WHEN TaxonName6.RankID = 20 THEN TaxonName6.TaxonName
		WHEN TaxonName7.RankID = 20 THEN TaxonName7.TaxonName
		WHEN TaxonName8.RankID = 20 THEN TaxonName5.TaxonName
		WHEN TaxonName9.RankID = 20 THEN TaxonName6.TaxonName
	END
	,	
	--SubkingdomAuthor
	CASE WHEN dbo.TaxonName.RankID = 20 THEN dbo.TaxonName.Author
		WHEN TaxonName1.RankID = 20 THEN TaxonName1.Author
		WHEN TaxonName2.RankID = 20 THEN TaxonName2.Author
		WHEN TaxonName3.RankID = 20 THEN TaxonName3.Author
		WHEN TaxonName4.RankID = 20 THEN TaxonName4.Author
		WHEN TaxonName5.RankID = 20 THEN TaxonName5.Author
		WHEN TaxonName6.RankID = 20 THEN TaxonName6.Author
		WHEN TaxonName7.RankID = 20 THEN TaxonName7.Author
		WHEN TaxonName8.RankID = 20 THEN TaxonName8.Author
		WHEN TaxonName9.RankID = 20 THEN TaxonName9.Author
	END
	,
	
	--3
	--Phyllum
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
	CASE WHEN dbo.TaxonName.RankID = 10 THEN dbo.TaxonName.Author
		WHEN TaxonName1.RankID = 10 THEN TaxonName1.Author
		WHEN TaxonName2.RankID = 10 THEN TaxonName2.Author
		WHEN TaxonName3.RankID = 10 THEN TaxonName3.Author
		WHEN TaxonName4.RankID = 10 THEN TaxonName4.Author
		WHEN TaxonName5.RankID = 10 THEN TaxonName5.Author
		WHEN TaxonName6.RankID = 10 THEN TaxonName6.Author
		WHEN TaxonName7.RankID = 10 THEN TaxonName7.Author
		WHEN TaxonName8.RankID = 10 THEN TaxonName8.Author
		WHEN TaxonName9.RankID = 10 THEN TaxonName9.Author
	END
	,
	--4
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
	
	--5
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
	--9
	--Tribe
	CASE WHEN dbo.TaxonName.RankID = 160 THEN dbo.TaxonName.TaxonName
		WHEN TaxonName1.RankID = 160 THEN TaxonName1.TaxonName
		WHEN TaxonName2.RankID = 160 THEN TaxonName2.TaxonName
		WHEN TaxonName3.RankID = 160 THEN TaxonName3.TaxonName
		WHEN TaxonName4.RankID = 160 THEN TaxonName4.TaxonName
		WHEN TaxonName5.RankID = 160 THEN TaxonName5.TaxonName
		WHEN TaxonName6.RankID = 160 THEN TaxonName6.TaxonName
		WHEN TaxonName7.RankID = 160 THEN TaxonName7.TaxonName
		WHEN TaxonName8.RankID = 160 THEN TaxonName8.TaxonName
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
		WHEN TaxonName7.RankID = 160 THEN TaxonName7.Author
		WHEN TaxonName8.RankID = 160 THEN TaxonName8.Author
	END
	,

	--10
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
	--11
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
	--12
	--Variety
	CASE WHEN dbo.TaxonName.RankID = 240 THEN dbo.TaxonName.TaxonName
		WHEN taxonName1.RankID = 240 THEN dbo.TaxonName.TaxonName
	END
	,
	--Variety Author
	CASE WHEN dbo.TaxonName.RankID = 240 THEN dbo.TaxonName.Author
		WHEN TaxonName1.RankID = 240 THEN dbo.TaxonName.Author
	END
	,
	dbo.Agent.FirstName,
	dbo.Agent.LastName,
	dbo.Determination.[date],
	dbo.Determination.TypeStatusName,
	dbo.Determination.[Current],
	dbo.Determination.Text1,
	dbo.Determination.Text2,
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
	att_Remarks	NTEXT,
	att_Text1	VARCHAR (300),
	att_Text2	VARCHAR (300),
	att_Text3	VARCHAR (300)

	)

INSERT INTO @Attributes
	(
	CollectionObject_ID3,
	att_Remarks,
	att_Text1,
	att_Text2,
	att_text3
	
)
SELECT

dbo.BiologicalObjectAttributes.BiologicalObjectAttributesID,
dbo.BiologicalObjectAttributes.Remarks,
dbo.BiologicalObjectAttributes.Text1,
dbo.BiologicalObjectAttributes.Text2,
dbo.BiologicalObjectAttributes.Text3

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

CASE WHEN  CollectionObject_1.PreparationMethod IS NULL THEN 'Dry'
	ELSE CollectionObject_1.PreparationMethod
	END
--	AS PreparationMethod
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
	PosRevcoNumber1,
	AndesTNumber2,
	AmbienteText1,
	CepaText2,
	CarMolecuRemarks,
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
	--1
	Kingdom1,
	KingdomAuthor1,
	--2
	Subkingdom1,
	SubkingdomAuthor1,
	--3	
	Phyllum1,
	PhyllumAuthor1,
	--4
	Class1,
	ClassAuthor1,
	--5
	Subclass1,
	SubclassAuthor1,
	--6
	Order1,
	OrderAuthor1,
	--7
	Suborder1,
	SuborderAuthor1,
	--8
	Family1,
	FamilyAuthor1,
	--9
	Tribe1,
	TribeAuthor1,
	--10
	Genus1,
	GenusAuthor1,
	--11
	Species1,
	SpeciesAuthor1,
	--12
	Variety1,
	VarietyAuthor1,

	--Determiner,
	DeterminerFirstName1,
	DeterminerLastName1,
	DeterminationDateSPY1,
	TypeStatusName1,
	Current1,
	Text1,
	Text2,
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
	CarMolecularRemarks,
	CarMicroText1,
	CarMicroText2,
	OtrosText3

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
	PosRevcoNumber_1,
	AndesTnumber_2,
	AmbienteText_1,
	CepaText_2,
	RemarksCollObj,
	'Separio',
	'Bacterias',
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
	 Subkingdom_Name,
	 Subkingdom_Author,
	 --3
	 Phyllum_name,
	 Phyllum_Author,
	 --4
	 Class_Name,
	 Class_Author,
	 --5
	 Subclass_Name,
	 Subclass_Author,
	 --6
	 Order_Name,
	 Order_Author,
	 --7
	 Suborder_Name,
	 Suborder_Author,
	 --8
	 Family_name,
	 Family_Author,
	 --9
	 Tribe_name,
	 Tribe_Author,
	 --10
	 Genus_Name,
	 Genus_Author,
	 --11
	 Species_Name,
	 Species_Author,
	 --12
	 Variety_Name,
	 Variety_Author,
	--Determiner,
	Determiner_FirstName,
	Determiner_LastName,
	[Determination_Date],
	TypeStatus_Name,
	isCurrent,
	Text_1,
	Text_2,
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
	att_Remarks,
	att_Text1,
	att_Text2,
	[att_Text3]

	
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