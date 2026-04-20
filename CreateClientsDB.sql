    CREATE TABLE Clients (
        Id          INT IDENTITY(1,1) PRIMARY KEY,
        ClientCode  NVARCHAR(20)  NOT NULL,
        FullName    NVARCHAR(150) NOT NULL,
        BirthDate   DATE          NOT NULL,
        PostalCode  NVARCHAR(10)  NULL,
        City        NVARCHAR(100) NULL,
        Street      NVARCHAR(100) NULL,
        House       NVARCHAR(10)  NULL,
        Apartment   NVARCHAR(10)  NULL,
        Email       NVARCHAR(200) NULL
    );