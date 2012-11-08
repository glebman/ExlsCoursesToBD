--select * from GradeBook where StudentID = 2276 AND RecordID IN (select RecordID from Schedule where CourseID = 94); 
--select * from Schedule where RecordID = 1574;
/*
select * from APersons
select * from Schedule where CourseID = 94 AND ClassID IN (Select GroupID from StudentsToGroups where StudentID = 2276 );
select * from GradeBook where StudentID = 2234;
select * from GradeBook where RecordID = 5100;
---select* from Groups where GroupID = 227;
--select* from GroupsToSchedule where GroupID = 227;
select * from StudentsToGroups where StudentID = 2276;
Select * from GradeTypes;
Select * from Groups
select * from GroupsToSchedule where GroupID
----------
Select CourseID from Courses where Name LIKE N'Алгоритмизация (VB/WSH)'; -- 380
Select PersonID from APersons where LastName Like N'Андреев' and FirstName Like N'Игорь' and MiddleName Like N'Владимирович'; --20384
*/

--------------то что надо!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
/*
select * from GradeBook where RecordID IN
(							
	select RecordID from Schedule where RecordID IN
	(
		select RecordID from GroupsToSchedule where GroupID IN
		(
			Select GroupID from StudentsToGroups where StudentID = (Select PersonID from APersons where LastName Like N'Жукова' and FirstName Like N'Анастасия' and MiddleName Like N'Алексеевна')
		)
		AND CourseID = (Select CourseID from Courses where Name LIKE N'Интерфейсы периферийных устройств')
	)
	
	AND StudentID = (Select PersonID from APersons where LastName Like N'Жукова' and FirstName Like N'Анастасия' and MiddleName Like N'Алексеевна')
	 
);
*/
-------------------record id последний по дате



/
DECLARE @lName nvarchar(20) = N'Жукова' ;
DECLARE @fName nvarchar(20) = N'Анастасия';
DECLARE @mName nvarchar(20) = N'Алексеевна';
DECLARE @year date = '1983-12-20'; 
DECLARE @course nvarchar(200) = N'Интерфейсы периферийных устройств';
DECLARE @mark INT = 77;
--DECLARE @record INT;
--DECLARE @record2 INT;
--DECLARE @studID INT;
DECLARE @gradeID INT;
--DECLARE @maxDate smalldatetime;
DECLARE @exam bit = 1;

EXEC GradeBook_InsertByFIO_YEAR_CourseName @lName,@fName,@mName,@year,@course,@mark,@exam
DECLARE @f INT;
Select @f=MAX([Key]) from GradeBook;	
select * from GradeBook where [key]= @f;


--==================================================================================================================
--------------------------------------------------------------------------------------------------------------------
---------------------------готовая процедура------------------------------------------------------------------------
--==================================================================================================================
DROP PROCEDURE GradeBook_InsertByFIO_YEAR_CourseName

CREATE PROCEDURE GradeBook_InsertByFIO_YEAR_CourseName
 @lName nvarchar(20),
 @fName nvarchar(20),
 @mName nvarchar(20),
 @year date,
 @course nvarchar(200),
 @mark INT,
 @exam bit,
 @gradeID INT 
AS
DECLARE  @studID INT;
DECLARE @maxDate smalldatetime;
DECLARE @record INT;
DECLARE @record2 INT;
--DECLARE @gradeID INT;
DECLARE @f INT;

Select @f=MAX([Key]) from GradeBook;
Select @studID =  PersonID from APersons where LastName Like @lName and FirstName Like @fName and MiddleName Like @mName and Birthday = @year;

select  TOP 1  @maxDate = StartDate,@record =  RecordID from Schedule where RecordID IN
	(
		select RecordID from GroupsToSchedule where GroupID IN
		(
			Select GroupID from StudentsToGroups where StudentID = @studID
		)
		AND CourseID = (Select CourseID from Courses where Name LIKE @course)
	)
	ORDER BY StartDate DESC;

IF @exam=0
BEGIN
select TOP 1 @record2=RecordID from Schedule where   StartDate < @maxDate AND RecordID IN
	(
		select RecordID from GroupsToSchedule where GroupID IN
		(
			Select GroupID from StudentsToGroups where StudentID = @studID
		)
		AND CourseID = (Select CourseID from Courses where Name LIKE @course)
	)
	ORDER BY StartDate DESC;
	
	INSERT INTO GradeBook  VALUES ((@f+1),@record2,@studID,@gradeID,5,@mark,NULL,NULL,NULL,NULL);
END	
ELSE
INSERT INTO GradeBook  VALUES ((@f+1),@record,@studID,@gradeID,4,@mark,NULL,NULL,NULL,NULL);

---------------------------------------------------------------------------------------------------
DECLARE @f INT;
Select @f=MAX([Key]) from GradeBook;	
PRINT @f;

PRINT (@f+1);


INSERT INTO GradeBook  VALUES ((@f+1),22,22,5,4,33,NULL,NULL,NULL,NULL)

select * from GradeBook where [key]=214878



---------добавление оценки
SET IDENTITY_INSERT GradeBook ON
INSERT INTO GradeBook  VALUES (NEXT,66298,20011,5,4,66,NULL,NULL,NULL,NULL)

INSERT INTO GradeBook (Key) VALUES (214873)

ALTER TABLE GradeBook AUTO_INCREMENT=214876

select * from GradeBook where RecordID = 66298

SELECT *
FROM sysindexes
WHERE id=OBJECT_ID('GradeBook')


--------------------------------------------------------------------------------------------------
--------------------------------------проверки---------------------------------------------------=
--------------------------------------------------------------------------------------------------

SELECT COUNT(*) FROM APersons WHERE  LastName Like N'Жукова' and FirstName Like N'Анастасия' and MiddleName Like N'Алексеевна' and Birthday = '1983-12-20';
SELECT * from GradeBook where GradeTypeID = 5


/*
			fName	"Игорь"	string
		lName	"Андреев"	string
		mName	"Владимирович"	string
		bYear	"1981-7-5"	string
		i	2	int
		subject	"Операционные системы и оболочки"	string
		mark	0	int
		procentMark	0	int
		exam	true	bool

*/