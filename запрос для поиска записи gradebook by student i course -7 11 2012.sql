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
CREATE PROCEDURE GradeBook_InsertByFIO_YEAR_CourseName


CREATE PROCEDURE GetRcordIDByFIO_YEAR_CourseName
DECLARE @lName nvarchar(20) = N'Жукова' ;
DECLARE @fName nvarchar(20) = N'Анастасия';
DECLARE @mName nvarchar(20) = N'Алексеевна';
DECLARE @year date = '1983-12-20'; 
DECLARE @course nvarchar(200) = N'Интерфейсы периферийных устройств';
DECLARE @mark INT;
DECLARE @record INT;

Select @record = RecordID from Schedule where StartDate = (
select MAX(StartDate)  from Schedule where RecordID IN
	(
		select RecordID from GroupsToSchedule where GroupID IN
		(
			Select GroupID from StudentsToGroups where StudentID = (Select PersonID from APersons where LastName Like @lName and FirstName Like @fName and MiddleName Like @mName and Birthday = @year)
		)
		AND CourseID = (Select CourseID from Courses where Name LIKE @course)
	)
)
And RecordID IN 
 (
 select RecordID  from Schedule where RecordID IN
	(
		select RecordID from GroupsToSchedule where GroupID IN
		(
			Select GroupID from StudentsToGroups where StudentID = (Select PersonID from APersons where LastName Like @lName and FirstName Like @fName and MiddleName Like @mName and Birthday = @year)
		)
		AND CourseID = (Select CourseID from Courses where Name LIKE @course)
	)
);



---------добавление оценки
SET IDENTITY_INSERT GradeBook ON
INSERT INTO GradeBook (Key, RecordID, StudentID,GradeID, GradeTypeID, Value) VALUES (214873,66298,20011,5,4,66)
ALTER TABLE GradeBook AUTO_INCREMENT=214873

select * from GradeBook