CREATE TABLE tblEntityScreening(
    EntityID INT IDENTITY(1,1) PRIMARY KEY,
    EntityName VARCHAR(128) UNIQUE NOT NULL,
    ChineseName NVARCHAR(128),
    RegCountry VARCHAR(128),
    CaseID VARCHAR(128),
    LastScreenDate DATETIME,
    ScreeningFlagPEP BIT,
    DetailsFlag NVARCHAR(4000),
    Archive bit default 0,
    ClientID int,
    FundID int
    CONSTRAINT UQ_EntityName_FundName_ClientName UNIQUE(EntityName, FundID, ClientID)
);
 
CREATE TABLE tblIndividualScreening(
    IndividualID INT IDENTITY(1,1) PRIMARY KEY,
    GivenName VARCHAR(128) NOT NULL,
    FamilyName VARCHAR(128) NOT NULL,
    ChineseName NVARCHAR(128),
    Gender VARCHAR(6),
    PlaceOfBirth VARCHAR(128),
    DateOfBirth VARCHAR(11),
    Citizenship VARCHAR(128),
    CityLocation VARCHAR(128),
    RegionLocation VARCHAR(128),
    CountryLocation VARCHAR(128),
    EntityID INT REFERENCES tblEntityScreening FOREIGN KEY (EntityID),
    Capacity VARCHAR(128),
    CaseID VARCHAR(128),
    LastScreenDate DATETIME,
    ScreeningFlagPEP BIT,
    DetailsFlag NVARCHAR(4000),
    Archive bit default 0,
    ClientID int,
    FundID int
    CONSTRAINT UQ_GivenName_FamilyName_FundName_ClientName UNIQUE(GivenName, FamilyName, FundID, ClientID)
);

-- make CaseID can have duplicate null
CREATE UNIQUE NONCLUSTERED INDEX idx_CaseID
ON tblIndividualScreening(CaseID)
WHERE CaseID IS NOT NULL;

CREATE UNIQUE NONCLUSTERED INDEX idx_CaseID
ON tblEntityScreening(CaseID)
WHERE CaseID IS NOT NULL;

/*
    Create new record in tblIndividualScreening
*/
CREATE PROCEDURE importIndividualScreening
    @GivenName VARCHAR(128),
    @FamilyName VARCHAR(128),
    @ChineseName NVARCHAR(128),
    @Gender VARCHAR(6),
    @PlaceOfBirth VARCHAR(128),
    @DateOfBirth VARCHAR(11),
    @Citizenship VARCHAR(128),
    @CityLocation VARCHAR(128),
    @RegionLocation VARCHAR(128),
    @CountryLocation VARCHAR(128),
    @EntityID INT,
    @FundID INT,
    @ClientID INT,
    @Capacity VARCHAR(128),
    @CaseID VARCHAR(128),
    @LastScreenDate VARCHAR(128),
    @PotentialPep Bit,
    @DetailsFlag VARCHAR(4000),
    @Archive Bit
AS
BEGIN
    DECLARE @InsertedCount INT = 0, @UpdatedCount INT = 0
    -- Check if a record with the same name exists
    IF EXISTS (SELECT 1 FROM tblIndividualScreening WHERE GivenName = @GivenName AND FamilyName = @FamilyName AND (FundID = @FundID OR (@FundID IS NULL AND FundID IS NULL)) AND (ClientID = @ClientID OR (@ClientID IS NULL AND ClientID IS NULL)))
    BEGIN
        -- If record exists, update it
        UPDATE tblIndividualScreening
        SET 
            ChineseName = @ChineseName,
            Gender = @Gender,
            PlaceOfBirth = @PlaceOfBirth,
            DateOfBirth = @DateOfBirth,
            Citizenship = @Citizenship,
            CityLocation = @CityLocation,
            RegionLocation = @RegionLocation,
            CountryLocation = @CountryLocation,
            EntityID = @EntityID,
            FundID = @FundID,
            ClientID = @ClientID,
            Capacity = @Capacity,
            CaseID = @CaseID,
            LastScreenDate = CAST(@LastScreenDate AS DATETIME),
            ScreeningFlagPep = @PotentialPep,
            DetailsFlag = @DetailsFlag,
            Archive = @Archive
        WHERE 
            GivenName = @GivenName AND FamilyName = @FamilyName AND (FundID = @FundID OR (@FundID IS NULL AND FundID IS NULL)) AND (ClientID = @ClientID OR (@ClientID IS NULL AND ClientID IS NULL))
        SET @UpdatedCount = @UpdatedCount + 1
    END
    ELSE
    BEGIN
        -- If record does not exist, insert a new one
        INSERT INTO tblIndividualScreening
        SELECT @GivenName, @FamilyName, @ChineseName, @Gender, @PlaceOfBirth, @DateOfBirth, @Citizenship, @CityLocation, @RegionLocation, @CountryLocation, @EntityID, @Capacity, @CaseID, CAST(@LastScreenDate AS DATETIME), @FundID, @ClientID, @PotentialPep, @DetailsFlag, @Archive
        Set @InsertedCount = @InsertedCount + 1
    END
    DECLARE @countTable Table(InsertCount INT, UpdateCount INT)
    INSERT INTO @countTable SELECT @InsertedCount, @UpdatedCount
    SELECT * FROM @countTable        
END;

/*
    Create new record in tblEntityScreening
*/
CREATE PROCEDURE importEntityScreening
    @EntityName VARCHAR(128),
    @ChineseName NVARCHAR(128),
    @RegCountry VARCHAR(128),
    @FundID INT,
    @ClientID INT,
    @CaseID VARCHAR(128),
    @LastScreenDate VARCHAR(128),
    @PotentialPep Bit,
    @DetailsFlag VARCHAR(4000),
    @Archive Bit
AS
BEGIN
    DECLARE @InsertedCount INT = 0, @UpdatedCount INT = 0
    -- Check if a record with the same EntityName exists
    IF EXISTS (SELECT 1 FROM tblEntityScreening WHERE EntityName = @EntityName AND (FundID = @FundID OR (@FundID IS NULL AND FundID IS NULL)) AND (ClientID = @ClientID OR (@ClientID IS NULL AND ClientID IS NULL)))
    BEGIN
        -- If record exists, update it
        UPDATE tblEntityScreening
        SET 
            ChineseName = @ChineseName, 
            RegCountry = @RegCountry, 
            CaseID = @CaseID, 
            LastScreenDate = CAST(@LastScreenDate AS DATETIME), 
            FundID = @FundID, 
            ClientID = @ClientID, 
            ScreeningFlagPep = @PotentialPep, 
            DetailsFlag = @DetailsFlag,
            Archive = @Archive
        WHERE
            EntityName = @EntityName AND (FundID = @FundID OR (@FundID IS NULL AND FundID IS NULL)) AND (ClientID = @ClientID OR (@ClientID IS NULL AND ClientID IS NULL))
        SET @UpdatedCount = @UpdatedCount + 1
    END
    ELSE
    BEGIN
        -- If record does not exist, insert a new one
        INSERT INTO tblEntityScreening
        SELECT @EntityName, @ChineseName, @RegCountry, @CaseID, CAST(@LastScreenDate AS DATETIME), @FundID, @ClientID, @PotentialPep, @DetailsFlag, @Archive
        Set @InsertedCount = @InsertedCount + 1
    END
    DECLARE @countTable Table(InsertCount INT, UpdateCount INT)
    INSERT INTO @countTable SELECT @InsertedCount, @UpdatedCount
    SELECT * FROM @countTable  
END;

/*
    Export records, filter by clientID
*/
CREATE FUNCTION exportScreeningWCByClient(@clientID Int)
RETURNS TABLE
AS
RETURN
Select i.GivenName + ' ' + i.FamilyName As Name, LEFT(i.CaseID + ' - ' + i.GivenName + ' ' + i.FamilyName, 50) As CaseID, 'IND' As EntityType, i.Gender, i.DateOfBirth, i.PlaceOfBirth, i.CountryLocation, i.Citizenship, '' As RegCountry, '' As IMOnumber, '' As IdentificationNumber
From tblIndividualScreening As i
Left Join tblEntityScreening As e On i.EntityID = e.EntityID
Where (@clientID = -1 Or Coalesce(i.ClientID, e.ClientID) = @clientID) And i.Archive != 1
Union
Select e.EntityName, LEFT(e.CaseID + ' - ' + e.EntityName, 50) As CaseID, 'ORG' As EntityType, '' As Gender, '' As DateOfBirth, '' As CountryLocation, '' As PlaceOfBirth, '' As Citizenship, e.RegCountry, '' As IMOnumber, '' As IdentificationNumber
From tblEntityScreening As e
Where (@clientID = -1 Or e.ClientID = @clientID) And e.Archive != 1;


/*
    Export records, filter by fundID
*/
CREATE FUNCTION exportScreeningWCByFund(@fundID INT)
RETURNS TABLE
AS
RETURN
Select i.GivenName + ' ' + i.FamilyName As Name, LEFT(i.CaseID + ' - ' + i.GivenName + ' ' + i.FamilyName, 50) As CaseID, 'IND' As EntityType, i.Gender, i.DateOfBirth, i.PlaceOfBirth, i.CountryLocation, i.Citizenship, '' As RegCountry, '' As IMOnumber, '' As IdentificationNumber
From tblIndividualScreening As i
Left Join tblEntityScreening As e On i.EntityID = e.EntityID
Where (@fundID = -1 Or Coalesce(i.FundID, e.FundID) = @fundID) And i.Archive != 1
Union
Select e.EntityName, LEFT(e.CaseID + ' - ' + e.EntityName, 50) As CaseID, 'ORG' As EntityType, '' As Gender, '' As DateOfBirth, '' As CountryLocation, '' As PlaceOfBirth, '' As Citizenship, e.RegCountry, '' As IMOnumber, '' As IdentificationNumber
From tblEntityScreening As e
Where (@fundID = -1 Or e.FundID = @fundID) And e.Archive != 1;