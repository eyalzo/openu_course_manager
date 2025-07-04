/****** Object:  Table [dbo].[CoursesNames]    Script Date: 27-11-01 11:54:55 ******/
CREATE TABLE [dbo].[CoursesNames] (
	[Course_Number] [numeric](18, 0) NOT NULL ,
	[Name] [nvarchar] (50) NOT NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[MamansErrors]    Script Date: 27-11-01 11:54:56 ******/
CREATE TABLE [dbo].[MamansErrors] (
	[MamanId] [int] NULL ,
	[QuestionNumber] [int] NULL ,
	[QuestionSection] [int] NULL ,
	[TextAnswer] [ntext] NULL ,
	[TextGuide] [ntext] NULL ,
	[Points] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

/****** Object:  Table [dbo].[Meetings]    Script Date: 27-11-01 11:54:56 ******/
CREATE TABLE [dbo].[Meetings] (
	[MeetingId] [int] NOT NULL ,
	[CourseId] [int] NOT NULL ,
	[Group] [int] NULL ,
	[Date] [smalldatetime] NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[Readings]    Script Date: 27-11-01 11:54:56 ******/
CREATE TABLE [dbo].[Readings] (
	[Reading_Id] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[Course_Number] [numeric](18, 0) NOT NULL ,
	[Description] [nvarchar] (30) NOT NULL ,
	[Detailed] [nvarchar] (200) NULL ,
	[Week_From] [numeric](18, 0) NULL ,
	[Week_To] [numeric](18, 0) NULL ,
	[Parent_Reading_Id] [numeric](18, 0) NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[Semesters]    Script Date: 27-11-01 11:54:57 ******/
CREATE TABLE [dbo].[Semesters] (
	[Semester] [nvarchar] (5) NOT NULL ,
	[Start_Date] [datetime] NOT NULL ,
	[Weeks] [numeric](18, 0) NOT NULL ,
	[Season] [nvarchar] (4) NOT NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[Students]    Script Date: 27-11-01 11:54:57 ******/
CREATE TABLE [dbo].[Students] (
	[Student_Id] [numeric](18, 0) NOT NULL ,
	[First] [nvarchar] (30) NOT NULL ,
	[Last] [nvarchar] (30) NOT NULL ,
	[Address] [nvarchar] (50) NULL ,
	[City] [nvarchar] (20) NULL ,
	[Phone_Mobile] [nvarchar] (20) NULL ,
	[Phone_Day] [nvarchar] (20) NULL ,
	[Phone_Evening] [nvarchar] (20) NULL ,
	[Last_Modified] [datetime] NOT NULL ,
	[Email] [varchar] (50) NULL ,
	[Code_Name] [nvarchar] (20) NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[StudyCenters]    Script Date: 27-11-01 11:54:57 ******/
CREATE TABLE [dbo].[StudyCenters] (
	[Center_Id] [numeric](18, 0) NOT NULL ,
	[Center_Name] [nvarchar] (50) NOT NULL ,
	[City] [nvarchar] (50) NOT NULL ,
	[Address] [nvarchar] (50) NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[Articles]    Script Date: 27-11-01 11:54:57 ******/
CREATE TABLE [dbo].[Articles] (
	[Article_Id] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[Name] [nvarchar] (100) NOT NULL ,
	[Basic_Type] [nvarchar] (20) NOT NULL ,
	[Course_Number] [numeric](18, 0) NOT NULL ,
	[Publish_Date] [datetime] NOT NULL ,
	[Author] [nvarchar] (50) NOT NULL ,
	[Description] [nvarchar] (200) NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[Courses]    Script Date: 27-11-01 11:54:57 ******/
CREATE TABLE [dbo].[Courses] (
	[Course_Id] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[Course_Number] [numeric](18, 0) NOT NULL ,
	[Semester] [nvarchar] (5) NOT NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[QuestionsText]    Script Date: 27-11-01 11:54:58 ******/
CREATE TABLE [dbo].[QuestionsText] (
	[Question_Id] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[Question_Text] [nvarchar] (2000) NULL ,
	[Answer_Text] [nvarchar] (2000) NULL ,
	[Comments_Text] [nvarchar] (1000) NULL ,
	[Date_Created] [datetime] NULL ,
	[Date_Answer_Modified] [datetime] NULL ,
	[Date_Question_Modified] [datetime] NULL ,
	[Entered_By] [numeric](18, 0) NULL ,
	[Reading_Id] [numeric](18, 0) NOT NULL ,
	[Source] [nvarchar] (100) NULL ,
	[Is_New] [nchar] (2) NULL ,
	[For_Exam] [nchar] (2) NULL ,
	[For_Maman] [nchar] (2) NULL ,
	[For_Guide] [nchar] (2) NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[CoursesGroups]    Script Date: 27-11-01 11:54:58 ******/
CREATE TABLE [dbo].[CoursesGroups] (
	[Group_Id] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[Course_Id] [numeric](18, 0) NOT NULL ,
	[Group_Number] [numeric](18, 0) NOT NULL ,
	[Center_Id] [numeric](18, 0) NOT NULL ,
	[Type_Mosadi] [nchar] (2) NULL ,
	[Type_Mugbar] [nchar] (2) NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[Exams]    Script Date: 27-11-01 11:54:58 ******/
CREATE TABLE [dbo].[Exams] (
	[Exam_Id] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[Course_Id] [numeric](18, 0) NOT NULL ,
	[Exam_Moed] [nvarchar] (2) NOT NULL ,
	[Moed_Symbol] [numeric](18, 0) NULL ,
	[Exam_Date] [datetime] NULL ,
	[Status] [nvarchar] (6) NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[Mamans]    Script Date: 27-11-01 11:54:58 ******/
CREATE TABLE [dbo].[Mamans] (
	[Maman_Id] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[Course_Id] [numeric](18, 0) NOT NULL ,
	[Maman_Number] [numeric](18, 0) NOT NULL ,
	[Delivery_Date] [datetime] NULL ,
	[Material] [nvarchar] (100) NULL ,
	[Weight] [numeric](18, 0) NULL ,
	[Mandatory] [nchar] (2) NULL ,
	[Max_Grade] [numeric](18, 0) NULL ,
	[Status] [nvarchar] (6) NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[QuestionsArticles]    Script Date: 27-11-01 11:54:59 ******/
CREATE TABLE [dbo].[QuestionsArticles] (
	[Article_Id] [numeric](18, 0) NOT NULL ,
	[Question_Id] [numeric](18, 0) NOT NULL ,
	[Question_Number] [varchar] (10) NULL ,
	[With_Solution] [bit] NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[QuestionsTextSub]    Script Date: 27-11-01 11:54:59 ******/
CREATE TABLE [dbo].[QuestionsTextSub] (
	[Question_Id] [numeric](18, 0) NOT NULL ,
	[Sub_Number] [numeric](18, 0) NOT NULL ,
	[Question_Text] [nvarchar] (1000) NULL ,
	[Answer_Text] [nvarchar] (1000) NULL ,
	[Relative_Grade] [numeric](18, 0) NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[ExamsQuestions]    Script Date: 27-11-01 11:54:59 ******/
CREATE TABLE [dbo].[ExamsQuestions] (
	[Exam_Id] [numeric](18, 0) NOT NULL ,
	[Question_Number] [numeric](18, 0) NOT NULL ,
	[Question_Id] [numeric](18, 0) NOT NULL ,
	[Max_Grade] [numeric](18, 0) NULL ,
	[Comments] [nvarchar] (1000) NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[MamansQuestions]    Script Date: 27-11-01 11:54:59 ******/
CREATE TABLE [dbo].[MamansQuestions] (
	[Question_Id] [numeric](18, 0) NOT NULL ,
	[Maman_Id] [numeric](18, 0) NOT NULL ,
	[Question_Number] [numeric](18, 0) NOT NULL ,
	[Max_Grade] [numeric](18, 0) NULL ,
	[Remarks] [nvarchar] (1000) NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[StudentsGroups]    Script Date: 27-11-01 11:54:59 ******/
CREATE TABLE [dbo].[StudentsGroups] (
	[Student_Id] [numeric](18, 0) NOT NULL ,
	[Group_Id] [numeric](18, 0) NOT NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[StudentsMamans]    Script Date: 27-11-01 11:54:59 ******/
CREATE TABLE [dbo].[StudentsMamans] (
	[Student_Id] [numeric](18, 0) NOT NULL ,
	[Maman_Id] [numeric](18, 0) NOT NULL ,
	[Date_Received] [datetime] NULL ,
	[Date_Sent] [datetime] NULL ,
	[Grade] [numeric](18, 0) NULL ,
	[Comments] [nvarchar] (200) NULL ,
	[Cancel_Late] [bit] NULL ,
	[Cancel_Copy] [bit] NULL ,
	[Last_Modified] [datetime] NULL ,
	[Second_Check] [nchar] (2) NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[StudentsMamansQuestions]    Script Date: 27-11-01 11:54:59 ******/
CREATE TABLE [dbo].[StudentsMamansQuestions] (
	[Student_Id] [numeric](18, 0) NOT NULL ,
	[Maman_Id] [numeric](18, 0) NOT NULL ,
	[Question_Number] [numeric](18, 0) NOT NULL ,
	[Grade] [numeric](18, 0) NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[CoursesNames] WITH NOCHECK ADD 
	CONSTRAINT [PK_CoursesNames] PRIMARY KEY  NONCLUSTERED 
	(
		[Course_Number]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Readings] WITH NOCHECK ADD 
	CONSTRAINT [PK_Material] PRIMARY KEY  NONCLUSTERED 
	(
		[Reading_Id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Semesters] WITH NOCHECK ADD 
	CONSTRAINT [PK_Semesters] PRIMARY KEY  NONCLUSTERED 
	(
		[Semester]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Students] WITH NOCHECK ADD 
	CONSTRAINT [DF_Students_Last_Modified] DEFAULT (getdate()) FOR [Last_Modified],
	CONSTRAINT [PK_Students] PRIMARY KEY  NONCLUSTERED 
	(
		[Student_Id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[StudyCenters] WITH NOCHECK ADD 
	CONSTRAINT [PK_StudyCenters] PRIMARY KEY  NONCLUSTERED 
	(
		[Center_Id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Articles] WITH NOCHECK ADD 
	CONSTRAINT [PK_Articles] PRIMARY KEY  NONCLUSTERED 
	(
		[Article_Id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Courses] WITH NOCHECK ADD 
	CONSTRAINT [PK_Courses] PRIMARY KEY  NONCLUSTERED 
	(
		[Course_Id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[QuestionsText] WITH NOCHECK ADD 
	CONSTRAINT [DF_QuestionsText_Date_Created] DEFAULT (getdate()) FOR [Date_Created],
	CONSTRAINT [DF_QuestionsText_Author] DEFAULT (29094661) FOR [Entered_By],
	CONSTRAINT [PK_QuestionsText] PRIMARY KEY  NONCLUSTERED 
	(
		[Question_Id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CoursesGroups] WITH NOCHECK ADD 
	CONSTRAINT [DF_CoursesGroups_Center_Id] DEFAULT (0) FOR [Center_Id],
	CONSTRAINT [DF_CoursesGroups_Group_Type] DEFAULT ('לא') FOR [Type_Mosadi],
	CONSTRAINT [DF_CoursesGroups_Type_Mugbar] DEFAULT ('לא') FOR [Type_Mugbar],
	CONSTRAINT [PK_CoursesGroups] PRIMARY KEY  NONCLUSTERED 
	(
		[Group_Id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Exams] WITH NOCHECK ADD 
	CONSTRAINT [DF_Exams_Status] DEFAULT ('בפיתוח') FOR [Status],
	CONSTRAINT [PK_Exams] PRIMARY KEY  NONCLUSTERED 
	(
		[Exam_Id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Mamans] WITH NOCHECK ADD 
	CONSTRAINT [DF_Mamans_Max_Grade] DEFAULT (100) FOR [Max_Grade],
	CONSTRAINT [DF_Mamans_Status] DEFAULT ('בפיתוח') FOR [Status],
	CONSTRAINT [PK_Mamans] PRIMARY KEY  NONCLUSTERED 
	(
		[Maman_Id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[QuestionsArticles] WITH NOCHECK ADD 
	CONSTRAINT [PK_QuestionsArticles] PRIMARY KEY  NONCLUSTERED 
	(
		[Article_Id],
		[Question_Id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ExamsQuestions] WITH NOCHECK ADD 
	CONSTRAINT [PK_ExamsQuestions] PRIMARY KEY  NONCLUSTERED 
	(
		[Exam_Id],
		[Question_Number]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MamansQuestions] WITH NOCHECK ADD 
	CONSTRAINT [PK_MamansQuestions] PRIMARY KEY  NONCLUSTERED 
	(
		[Maman_Id],
		[Question_Number]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[StudentsGroups] WITH NOCHECK ADD 
	CONSTRAINT [PK_StudentsGroups] PRIMARY KEY  NONCLUSTERED 
	(
		[Student_Id],
		[Group_Id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[StudentsMamans] WITH NOCHECK ADD 
	CONSTRAINT [DF_StudentsMamans_Cancel_Late] DEFAULT (0) FOR [Cancel_Late],
	CONSTRAINT [DF_StudentsMamans_Cancel_Copy] DEFAULT (0) FOR [Cancel_Copy],
	CONSTRAINT [DF_StudentsMamans_Last_Modified] DEFAULT (getdate()) FOR [Last_Modified],
	CONSTRAINT [PK_StudentsMamans] PRIMARY KEY  NONCLUSTERED 
	(
		[Student_Id],
		[Maman_Id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[StudentsMamansQuestions] WITH NOCHECK ADD 
	CONSTRAINT [PK_StudentsMamansQuestions] PRIMARY KEY  NONCLUSTERED 
	(
		[Student_Id],
		[Maman_Id],
		[Question_Number]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Articles] ADD 
	CONSTRAINT [FK_Articles_CoursesNames] FOREIGN KEY 
	(
		[Course_Number]
	) REFERENCES [dbo].[CoursesNames] (
		[Course_Number]
	) NOT FOR REPLICATION 
GO

ALTER TABLE [dbo].[Courses] ADD 
	CONSTRAINT [FK_Courses_CoursesNames] FOREIGN KEY 
	(
		[Course_Number]
	) REFERENCES [dbo].[CoursesNames] (
		[Course_Number]
	) NOT FOR REPLICATION ,
	CONSTRAINT [FK_Courses_Semesters] FOREIGN KEY 
	(
		[Semester]
	) REFERENCES [dbo].[Semesters] (
		[Semester]
	) NOT FOR REPLICATION 
GO

ALTER TABLE [dbo].[QuestionsText] ADD 
	CONSTRAINT [FK_QuestionsText_Material] FOREIGN KEY 
	(
		[Reading_Id]
	) REFERENCES [dbo].[Readings] (
		[Reading_Id]
	) NOT FOR REPLICATION 
GO

ALTER TABLE [dbo].[CoursesGroups] ADD 
	CONSTRAINT [FK_CoursesGroups_Courses] FOREIGN KEY 
	(
		[Course_Id]
	) REFERENCES [dbo].[Courses] (
		[Course_Id]
	) NOT FOR REPLICATION 
GO

ALTER TABLE [dbo].[Exams] ADD 
	CONSTRAINT [FK_Exams_Courses] FOREIGN KEY 
	(
		[Course_Id]
	) REFERENCES [dbo].[Courses] (
		[Course_Id]
	) NOT FOR REPLICATION 
GO

ALTER TABLE [dbo].[Mamans] ADD 
	CONSTRAINT [FK_Mamans_Courses] FOREIGN KEY 
	(
		[Course_Id]
	) REFERENCES [dbo].[Courses] (
		[Course_Id]
	) NOT FOR REPLICATION 
GO

ALTER TABLE [dbo].[QuestionsArticles] ADD 
	CONSTRAINT [FK_QuestionsArticles_Articles] FOREIGN KEY 
	(
		[Article_Id]
	) REFERENCES [dbo].[Articles] (
		[Article_Id]
	) NOT FOR REPLICATION ,
	CONSTRAINT [FK_QuestionsArticles_QuestionsText] FOREIGN KEY 
	(
		[Question_Id]
	) REFERENCES [dbo].[QuestionsText] (
		[Question_Id]
	) NOT FOR REPLICATION 
GO

ALTER TABLE [dbo].[QuestionsTextSub] ADD 
	CONSTRAINT [FK_QuestionsTextSub_QuestionsText] FOREIGN KEY 
	(
		[Question_Id]
	) REFERENCES [dbo].[QuestionsText] (
		[Question_Id]
	) NOT FOR REPLICATION 
GO

ALTER TABLE [dbo].[ExamsQuestions] ADD 
	CONSTRAINT [FK_ExamsQuestions_Exams] FOREIGN KEY 
	(
		[Exam_Id]
	) REFERENCES [dbo].[Exams] (
		[Exam_Id]
	) NOT FOR REPLICATION ,
	CONSTRAINT [FK_ExamsQuestions_QuestionsText] FOREIGN KEY 
	(
		[Question_Id]
	) REFERENCES [dbo].[QuestionsText] (
		[Question_Id]
	) NOT FOR REPLICATION 
GO

ALTER TABLE [dbo].[MamansQuestions] ADD 
	CONSTRAINT [FK_MamansQuestions_Mamans] FOREIGN KEY 
	(
		[Maman_Id]
	) REFERENCES [dbo].[Mamans] (
		[Maman_Id]
	) NOT FOR REPLICATION ,
	CONSTRAINT [FK_MamansQuestions_QuestionsText] FOREIGN KEY 
	(
		[Question_Id]
	) REFERENCES [dbo].[QuestionsText] (
		[Question_Id]
	) NOT FOR REPLICATION 
GO

ALTER TABLE [dbo].[StudentsGroups] ADD 
	CONSTRAINT [FK_StudentsGroups_CoursesGroups] FOREIGN KEY 
	(
		[Group_Id]
	) REFERENCES [dbo].[CoursesGroups] (
		[Group_Id]
	) NOT FOR REPLICATION ,
	CONSTRAINT [FK_StudentsGroups_Students] FOREIGN KEY 
	(
		[Student_Id]
	) REFERENCES [dbo].[Students] (
		[Student_Id]
	) NOT FOR REPLICATION 
GO

ALTER TABLE [dbo].[StudentsMamans] ADD 
	CONSTRAINT [FK_StudentsMamans_Mamans] FOREIGN KEY 
	(
		[Maman_Id]
	) REFERENCES [dbo].[Mamans] (
		[Maman_Id]
	) NOT FOR REPLICATION ,
	CONSTRAINT [FK_StudentsMamans_Students] FOREIGN KEY 
	(
		[Student_Id]
	) REFERENCES [dbo].[Students] (
		[Student_Id]
	) NOT FOR REPLICATION 
GO

ALTER TABLE [dbo].[StudentsMamansQuestions] ADD 
	CONSTRAINT [FK_StudentsMamansQuestions_StudentsMamans] FOREIGN KEY 
	(
		[Student_Id],
		[Maman_Id]
	) REFERENCES [dbo].[StudentsMamans] (
		[Student_Id],
		[Maman_Id]
	) NOT FOR REPLICATION 
GO

