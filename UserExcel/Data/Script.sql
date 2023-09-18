
 CREATE DATABASE UserDB

 Create Table UserMst(
 Id int Not Null Identity(1,1) Primary Key,
 FirstName [Varchar](50) Not Null,
 MiddleName [Varchar](50) Not Null,
 LastName [Varchar](50) Not Null,
 UserName [Varchar](50) Not Null,
 Password [Varchar](50) Not Null,
 Address [Varchar](100) Not Null,
 Pincode [Varchar](50) Not Null,
 Mobile1 [Varchar](50) Not Null,
 Mobile2 [Varchar](50) Null,
 Email [Varchar](50) Not Null,
 CompanyName [Varchar](50) Not Null,
 IsActive [Bit] Not Null,
 IsDelete [Bit] Not Null,
 UpdateDate [datetime] Null,
 CreateDate [datetime] Null
)


 select *from UserMst
