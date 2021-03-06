USE [master]
GO
/****** Object:  Database [LMBA]    Script Date: 23/05/2022 9:51:12 ******/
CREATE DATABASE [LMBA]
 CONTAINMENT = NONE
 ON  PRIMARY 
( NAME = N'LMBA', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL15.MSSQLSERVER\MSSQL\DATA\LMBA.mdf' , SIZE = 8192KB , MAXSIZE = UNLIMITED, FILEGROWTH = 65536KB )
 LOG ON 
( NAME = N'LMBA_log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL15.MSSQLSERVER\MSSQL\DATA\LMBA_log.ldf' , SIZE = 8192KB , MAXSIZE = 2048GB , FILEGROWTH = 65536KB )
 WITH CATALOG_COLLATION = DATABASE_DEFAULT
GO
ALTER DATABASE [LMBA] SET COMPATIBILITY_LEVEL = 150
GO
IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [LMBA].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO
ALTER DATABASE [LMBA] SET ANSI_NULL_DEFAULT OFF 
GO
ALTER DATABASE [LMBA] SET ANSI_NULLS OFF 
GO
ALTER DATABASE [LMBA] SET ANSI_PADDING OFF 
GO
ALTER DATABASE [LMBA] SET ANSI_WARNINGS OFF 
GO
ALTER DATABASE [LMBA] SET ARITHABORT OFF 
GO
ALTER DATABASE [LMBA] SET AUTO_CLOSE OFF 
GO
ALTER DATABASE [LMBA] SET AUTO_SHRINK OFF 
GO
ALTER DATABASE [LMBA] SET AUTO_UPDATE_STATISTICS ON 
GO
ALTER DATABASE [LMBA] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO
ALTER DATABASE [LMBA] SET CURSOR_DEFAULT  GLOBAL 
GO
ALTER DATABASE [LMBA] SET CONCAT_NULL_YIELDS_NULL OFF 
GO
ALTER DATABASE [LMBA] SET NUMERIC_ROUNDABORT OFF 
GO
ALTER DATABASE [LMBA] SET QUOTED_IDENTIFIER OFF 
GO
ALTER DATABASE [LMBA] SET RECURSIVE_TRIGGERS OFF 
GO
ALTER DATABASE [LMBA] SET  DISABLE_BROKER 
GO
ALTER DATABASE [LMBA] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO
ALTER DATABASE [LMBA] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO
ALTER DATABASE [LMBA] SET TRUSTWORTHY OFF 
GO
ALTER DATABASE [LMBA] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO
ALTER DATABASE [LMBA] SET PARAMETERIZATION SIMPLE 
GO
ALTER DATABASE [LMBA] SET READ_COMMITTED_SNAPSHOT OFF 
GO
ALTER DATABASE [LMBA] SET HONOR_BROKER_PRIORITY OFF 
GO
ALTER DATABASE [LMBA] SET RECOVERY SIMPLE 
GO
ALTER DATABASE [LMBA] SET  MULTI_USER 
GO
ALTER DATABASE [LMBA] SET PAGE_VERIFY CHECKSUM  
GO
ALTER DATABASE [LMBA] SET DB_CHAINING OFF 
GO
ALTER DATABASE [LMBA] SET FILESTREAM( NON_TRANSACTED_ACCESS = OFF ) 
GO
ALTER DATABASE [LMBA] SET TARGET_RECOVERY_TIME = 60 SECONDS 
GO
ALTER DATABASE [LMBA] SET DELAYED_DURABILITY = DISABLED 
GO
ALTER DATABASE [LMBA] SET ACCELERATED_DATABASE_RECOVERY = OFF  
GO
ALTER DATABASE [LMBA] SET QUERY_STORE = OFF
GO
USE [LMBA]
GO
/****** Object:  Table [dbo].[Cargo]    Script Date: 23/05/2022 9:51:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Cargo](
	[idCargo] [int] IDENTITY(1,1) NOT NULL,
	[nombreCargo] [nvarchar](30) NOT NULL,
	[descripcion] [nvarchar](30) NOT NULL,
	[estado] [bit] NOT NULL,
	[idDepartamento] [int] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Departamento]    Script Date: 23/05/2022 9:51:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Departamento](
	[idDepartamento] [int] IDENTITY(1,1) NOT NULL,
	[nombreDepartamento] [nvarchar](30) NOT NULL,
	[ext] [nvarchar](10) NOT NULL,
	[email] [nvarchar](100) NULL,
	[estado] [int] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Empleado]    Script Date: 23/05/2022 9:51:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Empleado](
	[idEmpleado] [int] IDENTITY(1,1) NOT NULL,
	[numCedula] [nvarchar](16) NOT NULL,
	[estado] [int] NOT NULL,
	[primerNombre] [nvarchar](30) NOT NULL,
	[segundoNombre] [nvarchar](10) NULL,
	[primerApellido] [nvarchar](30) NOT NULL,
	[segundoApellido] [nvarchar](30) NULL,
	[fechaNacimiento] [datetime] NOT NULL,
	[sexo] [bit] NOT NULL,
	[fechaIngreso] [datetime] NULL,
	[direccion] [nvarchar](100) NOT NULL,
	[observacion] [nvarchar](100) NULL,
	[fotoEmpleado] [image] NULL,
	[telefono] [nvarchar](10) NULL,
	[emailPersonal] [nvarchar](100) NULL,
	[emailCorporativo] [nvarchar](100) NULL,
	[idCargo] [int] NOT NULL,
	[idHorario] [int] NOT NULL,
	[id_user] [int] NULL,
	[PIN] [nvarchar](4) NOT NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[empleadoRegistro]    Script Date: 23/05/2022 9:51:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[empleadoRegistro](
	[idEmpleadoRegistro] [int] IDENTITY(1,1) NOT NULL,
	[idRegistro] [int] NOT NULL,
	[idEmpleado] [int] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Horario]    Script Date: 23/05/2022 9:51:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Horario](
	[idHorario] [int] IDENTITY(1,1) NOT NULL,
	[nombre] [nvarchar](30) NOT NULL,
	[horaEntrada] [time](7) NOT NULL,
	[horaSalida] [time](7) NOT NULL,
	[almuerzo] [time](7) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Justificacion]    Script Date: 23/05/2022 9:51:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Justificacion](
	[idJustificacion] [int] IDENTITY(1,1) NOT NULL,
	[estado] [int] NULL,
	[descripcion] [nvarchar](100) NULL,
	[fechaEntrada] [datetime] NOT NULL,
	[fechaSalida] [datetime] NOT NULL,
	[horaEntrada] [time](7) NOT NULL,
	[horaSalida] [time](7) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[registroES]    Script Date: 23/05/2022 9:51:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[registroES](
	[idRegistro] [int] IDENTITY(1,1) NOT NULL,
	[estado] [int] NOT NULL,
	[fecha] [datetime] NULL,
	[horaEntrada] [datetime] NULL,
	[horaSalida] [datetime] NULL,
	[idJustificacion] [int] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[tbl_opcion]    Script Date: 23/05/2022 9:51:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_opcion](
	[id_opcion] [int] IDENTITY(1,1) NOT NULL,
	[opcion] [varchar](50) NOT NULL,
	[estado] [int] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[tbl_rol]    Script Date: 23/05/2022 9:51:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_rol](
	[id_rol] [int] IDENTITY(1,1) NOT NULL,
	[rol] [varchar](50) NOT NULL,
	[estado] [varchar](50) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[tbl_rolOpcion]    Script Date: 23/05/2022 9:51:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_rolOpcion](
	[id_rolOpcion] [int] IDENTITY(1,1) NOT NULL,
	[id_opcion] [int] NOT NULL,
	[id_rol] [int] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[tbl_user]    Script Date: 23/05/2022 9:51:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_user](
	[id_user] [int] IDENTITY(1,1) NOT NULL,
	[username] [varchar](50) NOT NULL,
	[pwd] [varchar](50) NOT NULL,
	[nombres] [varchar](50) NOT NULL,
	[apellidos] [varchar](50) NOT NULL,
	[pwd_temp] [varchar](50) NULL,
	[email] [varchar](50) NOT NULL,
	[estado] [int] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[tbl_userRol]    Script Date: 23/05/2022 9:51:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_userRol](
	[id_UserRol] [int] IDENTITY(1,1) NOT NULL,
	[id_rol] [int] NOT NULL,
	[id_user] [int] NOT NULL
) ON [PRIMARY]
GO
SET IDENTITY_INSERT [dbo].[Cargo] ON 

INSERT [dbo].[Cargo] ([idCargo], [nombreCargo], [descripcion], [estado], [idDepartamento]) VALUES (1, N'Analista de Sistemas', N'Analiza Sistemas', 1, 1)
INSERT [dbo].[Cargo] ([idCargo], [nombreCargo], [descripcion], [estado], [idDepartamento]) VALUES (2, N'Programador', N'Programa', 1, 1)
INSERT [dbo].[Cargo] ([idCargo], [nombreCargo], [descripcion], [estado], [idDepartamento]) VALUES (3, N'Programadora', N'También programa', 1, 1)
INSERT [dbo].[Cargo] ([idCargo], [nombreCargo], [descripcion], [estado], [idDepartamento]) VALUES (4, N'Diseñador', N'Diseña cosas', 1, 2)
INSERT [dbo].[Cargo] ([idCargo], [nombreCargo], [descripcion], [estado], [idDepartamento]) VALUES (5, N'Reclutador', N'Recluta gente', 0, 3)
SET IDENTITY_INSERT [dbo].[Cargo] OFF
GO
SET IDENTITY_INSERT [dbo].[Departamento] ON 

INSERT [dbo].[Departamento] ([idDepartamento], [nombreDepartamento], [ext], [email], [estado]) VALUES (1, N'Desarrollo', N'19', N'desarrollo@lmba.com', 1)
INSERT [dbo].[Departamento] ([idDepartamento], [nombreDepartamento], [ext], [email], [estado]) VALUES (2, N'Diseño', N'20', N'diseño@lmba.com', 1)
INSERT [dbo].[Departamento] ([idDepartamento], [nombreDepartamento], [ext], [email], [estado]) VALUES (3, N'RRHH', N'18', N'rrhh@lmba.com', 1)
SET IDENTITY_INSERT [dbo].[Departamento] OFF
GO
SET IDENTITY_INSERT [dbo].[Empleado] ON 

INSERT [dbo].[Empleado] ([idEmpleado], [numCedula], [estado], [primerNombre], [segundoNombre], [primerApellido], [segundoApellido], [fechaNacimiento], [sexo], [fechaIngreso], [direccion], [observacion], [fotoEmpleado], [telefono], [emailPersonal], [emailCorporativo], [idCargo], [idHorario], [id_user], [PIN]) VALUES (8, N'001-190101-1009G', 1, N'Armando', N'Alexander', N'Meza', N'López', CAST(N'2001-01-19T00:00:00.000' AS DateTime), 1, CAST(N'2022-05-21T00:00:00.000' AS DateTime), N'De la clínica Don Bosco, 4c. al este, 2c. al norte.', N'f', NULL, N'77665604', N'armando@gmail.com', N'armando@lmba.com', 1, 1, NULL, N'1234')
INSERT [dbo].[Empleado] ([idEmpleado], [numCedula], [estado], [primerNombre], [segundoNombre], [primerApellido], [segundoApellido], [fechaNacimiento], [sexo], [fechaIngreso], [direccion], [observacion], [fotoEmpleado], [telefono], [emailPersonal], [emailCorporativo], [idCargo], [idHorario], [id_user], [PIN]) VALUES (9, N'041-080804-1000B', 1, N'David', N'Sebastián', N'Parrales', N'Ponce', CAST(N'2004-08-08T00:00:00.000' AS DateTime), 1, CAST(N'2022-05-21T00:00:00.000' AS DateTime), N'De la Ponce 1 arriba, 1 abajo', N'', NULL, N'88888888', N'ponce@gmail.com', N'ponce@lmba.com', 2, 1, NULL, N'1234')
INSERT [dbo].[Empleado] ([idEmpleado], [numCedula], [estado], [primerNombre], [segundoNombre], [primerApellido], [segundoApellido], [fechaNacimiento], [sexo], [fechaIngreso], [direccion], [observacion], [fotoEmpleado], [telefono], [emailPersonal], [emailCorporativo], [idCargo], [idHorario], [id_user], [PIN]) VALUES (13, N'001-151001-8888A', 1, N'Alejandra', N'', N'Prado', N'Sanchez', CAST(N'2022-05-22T00:00:00.000' AS DateTime), 0, CAST(N'2022-05-22T00:00:00.000' AS DateTime), N'Tipitapa', N'', NULL, N'11112222', N'ale@gmail.com', N'ale@lmba.com', 4, 1, NULL, N'1234')
SET IDENTITY_INSERT [dbo].[Empleado] OFF
GO
SET IDENTITY_INSERT [dbo].[empleadoRegistro] ON 

INSERT [dbo].[empleadoRegistro] ([idEmpleadoRegistro], [idRegistro], [idEmpleado]) VALUES (3, 1, 8)
INSERT [dbo].[empleadoRegistro] ([idEmpleadoRegistro], [idRegistro], [idEmpleado]) VALUES (6, 2, 8)
SET IDENTITY_INSERT [dbo].[empleadoRegistro] OFF
GO
SET IDENTITY_INSERT [dbo].[Horario] ON 

INSERT [dbo].[Horario] ([idHorario], [nombre], [horaEntrada], [horaSalida], [almuerzo]) VALUES (1, N'regular', CAST(N'08:00:00' AS Time), CAST(N'17:00:00' AS Time), CAST(N'01:00:00' AS Time))
SET IDENTITY_INSERT [dbo].[Horario] OFF
GO
SET IDENTITY_INSERT [dbo].[Justificacion] ON 

INSERT [dbo].[Justificacion] ([idJustificacion], [estado], [descripcion], [fechaEntrada], [fechaSalida], [horaEntrada], [horaSalida]) VALUES (1, 1, N'Enfermedad mental', CAST(N'2022-05-23T00:00:00.000' AS DateTime), CAST(N'2022-05-23T00:00:00.000' AS DateTime), CAST(N'08:00:00' AS Time), CAST(N'17:00:00' AS Time))
INSERT [dbo].[Justificacion] ([idJustificacion], [estado], [descripcion], [fechaEntrada], [fechaSalida], [horaEntrada], [horaSalida]) VALUES (2, 1, N'Enfermedad pulmonar', CAST(N'2022-05-23T00:00:00.000' AS DateTime), CAST(N'2022-05-23T00:00:00.000' AS DateTime), CAST(N'07:55:00' AS Time), CAST(N'17:05:00' AS Time))
SET IDENTITY_INSERT [dbo].[Justificacion] OFF
GO
SET IDENTITY_INSERT [dbo].[registroES] ON 

INSERT [dbo].[registroES] ([idRegistro], [estado], [fecha], [horaEntrada], [horaSalida], [idJustificacion]) VALUES (1, 1, CAST(N'2022-05-22T00:00:00.000' AS DateTime), CAST(N'2022-05-22T08:00:00.000' AS DateTime), CAST(N'2022-05-22T17:05:00.000' AS DateTime), NULL)
INSERT [dbo].[registroES] ([idRegistro], [estado], [fecha], [horaEntrada], [horaSalida], [idJustificacion]) VALUES (2, 1, CAST(N'2022-05-23T00:00:00.000' AS DateTime), CAST(N'2022-05-23T07:30:00.000' AS DateTime), CAST(N'2022-05-23T14:00:00.000' AS DateTime), NULL)
SET IDENTITY_INSERT [dbo].[registroES] OFF
GO
SET IDENTITY_INSERT [dbo].[tbl_rol] ON 

INSERT [dbo].[tbl_rol] ([id_rol], [rol], [estado]) VALUES (2, N'admin', N'1')
SET IDENTITY_INSERT [dbo].[tbl_rol] OFF
GO
SET IDENTITY_INSERT [dbo].[tbl_user] ON 

INSERT [dbo].[tbl_user] ([id_user], [username], [pwd], [nombres], [apellidos], [pwd_temp], [email], [estado]) VALUES (2, N'armando', N'123', N'Armando Alexander', N'Meza Lopez', N'123', N'armando@lmba.com', 1)
SET IDENTITY_INSERT [dbo].[tbl_user] OFF
GO
SET IDENTITY_INSERT [dbo].[tbl_userRol] ON 

INSERT [dbo].[tbl_userRol] ([id_UserRol], [id_rol], [id_user]) VALUES (1, 2, 2)
SET IDENTITY_INSERT [dbo].[tbl_userRol] OFF
GO
/****** Object:  Index [PK2]    Script Date: 23/05/2022 9:51:12 ******/
ALTER TABLE [dbo].[Cargo] ADD  CONSTRAINT [PK2] PRIMARY KEY NONCLUSTERED 
(
	[idCargo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
/****** Object:  Index [PK3]    Script Date: 23/05/2022 9:51:12 ******/
ALTER TABLE [dbo].[Departamento] ADD  CONSTRAINT [PK3] PRIMARY KEY NONCLUSTERED 
(
	[idDepartamento] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
/****** Object:  Index [PK1]    Script Date: 23/05/2022 9:51:12 ******/
ALTER TABLE [dbo].[Empleado] ADD  CONSTRAINT [PK1] PRIMARY KEY NONCLUSTERED 
(
	[idEmpleado] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
/****** Object:  Index [PK3_1]    Script Date: 23/05/2022 9:51:12 ******/
ALTER TABLE [dbo].[empleadoRegistro] ADD  CONSTRAINT [PK3_1] PRIMARY KEY NONCLUSTERED 
(
	[idEmpleadoRegistro] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
/****** Object:  Index [PK22]    Script Date: 23/05/2022 9:51:12 ******/
ALTER TABLE [dbo].[Horario] ADD  CONSTRAINT [PK22] PRIMARY KEY NONCLUSTERED 
(
	[idHorario] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
/****** Object:  Index [PK20]    Script Date: 23/05/2022 9:51:12 ******/
ALTER TABLE [dbo].[Justificacion] ADD  CONSTRAINT [PK20] PRIMARY KEY NONCLUSTERED 
(
	[idJustificacion] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
/****** Object:  Index [PK2_1]    Script Date: 23/05/2022 9:51:12 ******/
ALTER TABLE [dbo].[registroES] ADD  CONSTRAINT [PK2_1] PRIMARY KEY NONCLUSTERED 
(
	[idRegistro] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
/****** Object:  Index [PK7]    Script Date: 23/05/2022 9:51:12 ******/
ALTER TABLE [dbo].[tbl_opcion] ADD  CONSTRAINT [PK7] PRIMARY KEY NONCLUSTERED 
(
	[id_opcion] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
/****** Object:  Index [PK12]    Script Date: 23/05/2022 9:51:12 ******/
ALTER TABLE [dbo].[tbl_rol] ADD  CONSTRAINT [PK12] PRIMARY KEY NONCLUSTERED 
(
	[id_rol] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
/****** Object:  Index [PK11]    Script Date: 23/05/2022 9:51:12 ******/
ALTER TABLE [dbo].[tbl_rolOpcion] ADD  CONSTRAINT [PK11] PRIMARY KEY NONCLUSTERED 
(
	[id_rolOpcion] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
/****** Object:  Index [PK8]    Script Date: 23/05/2022 9:51:12 ******/
ALTER TABLE [dbo].[tbl_user] ADD  CONSTRAINT [PK8] PRIMARY KEY NONCLUSTERED 
(
	[id_user] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
/****** Object:  Index [PK9]    Script Date: 23/05/2022 9:51:12 ******/
ALTER TABLE [dbo].[tbl_userRol] ADD  CONSTRAINT [PK9] PRIMARY KEY NONCLUSTERED 
(
	[id_UserRol] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
ALTER TABLE [dbo].[Cargo] ADD  DEFAULT ((1)) FOR [estado]
GO
ALTER TABLE [dbo].[Departamento] ADD  DEFAULT ((1)) FOR [estado]
GO
ALTER TABLE [dbo].[Empleado] ADD  CONSTRAINT [DF__Empleado__estado__2A4B4B5E]  DEFAULT ((1)) FOR [estado]
GO
ALTER TABLE [dbo].[Empleado] ADD  CONSTRAINT [DF__Empleado__sexo__2B3F6F97]  DEFAULT ((1)) FOR [sexo]
GO
ALTER TABLE [dbo].[Justificacion] ADD  CONSTRAINT [DF__Justifica__estad__31EC6D26]  DEFAULT ((1)) FOR [estado]
GO
ALTER TABLE [dbo].[registroES] ADD  DEFAULT ((1)) FOR [estado]
GO
ALTER TABLE [dbo].[tbl_opcion] ADD  DEFAULT ((1)) FOR [estado]
GO
ALTER TABLE [dbo].[Cargo]  WITH CHECK ADD  CONSTRAINT [RefDepartamento1] FOREIGN KEY([idDepartamento])
REFERENCES [dbo].[Departamento] ([idDepartamento])
GO
ALTER TABLE [dbo].[Cargo] CHECK CONSTRAINT [RefDepartamento1]
GO
ALTER TABLE [dbo].[Empleado]  WITH CHECK ADD  CONSTRAINT [RefCargo16] FOREIGN KEY([idCargo])
REFERENCES [dbo].[Cargo] ([idCargo])
GO
ALTER TABLE [dbo].[Empleado] CHECK CONSTRAINT [RefCargo16]
GO
ALTER TABLE [dbo].[Empleado]  WITH CHECK ADD  CONSTRAINT [RefHorario17] FOREIGN KEY([idHorario])
REFERENCES [dbo].[Horario] ([idHorario])
GO
ALTER TABLE [dbo].[Empleado] CHECK CONSTRAINT [RefHorario17]
GO
ALTER TABLE [dbo].[Empleado]  WITH CHECK ADD  CONSTRAINT [Reftbl_user18] FOREIGN KEY([id_user])
REFERENCES [dbo].[tbl_user] ([id_user])
GO
ALTER TABLE [dbo].[Empleado] CHECK CONSTRAINT [Reftbl_user18]
GO
ALTER TABLE [dbo].[empleadoRegistro]  WITH CHECK ADD  CONSTRAINT [RefEmpleado13] FOREIGN KEY([idEmpleado])
REFERENCES [dbo].[Empleado] ([idEmpleado])
GO
ALTER TABLE [dbo].[empleadoRegistro] CHECK CONSTRAINT [RefEmpleado13]
GO
ALTER TABLE [dbo].[empleadoRegistro]  WITH CHECK ADD  CONSTRAINT [RefregistroES12] FOREIGN KEY([idRegistro])
REFERENCES [dbo].[registroES] ([idRegistro])
GO
ALTER TABLE [dbo].[empleadoRegistro] CHECK CONSTRAINT [RefregistroES12]
GO
ALTER TABLE [dbo].[registroES]  WITH CHECK ADD  CONSTRAINT [RefJustificacion14] FOREIGN KEY([idJustificacion])
REFERENCES [dbo].[Justificacion] ([idJustificacion])
GO
ALTER TABLE [dbo].[registroES] CHECK CONSTRAINT [RefJustificacion14]
GO
ALTER TABLE [dbo].[tbl_rolOpcion]  WITH CHECK ADD  CONSTRAINT [Reftbl_opcion8] FOREIGN KEY([id_opcion])
REFERENCES [dbo].[tbl_opcion] ([id_opcion])
GO
ALTER TABLE [dbo].[tbl_rolOpcion] CHECK CONSTRAINT [Reftbl_opcion8]
GO
ALTER TABLE [dbo].[tbl_rolOpcion]  WITH CHECK ADD  CONSTRAINT [Reftbl_rol9] FOREIGN KEY([id_rol])
REFERENCES [dbo].[tbl_rol] ([id_rol])
GO
ALTER TABLE [dbo].[tbl_rolOpcion] CHECK CONSTRAINT [Reftbl_rol9]
GO
ALTER TABLE [dbo].[tbl_userRol]  WITH CHECK ADD  CONSTRAINT [Reftbl_rol10] FOREIGN KEY([id_rol])
REFERENCES [dbo].[tbl_rol] ([id_rol])
GO
ALTER TABLE [dbo].[tbl_userRol] CHECK CONSTRAINT [Reftbl_rol10]
GO
ALTER TABLE [dbo].[tbl_userRol]  WITH CHECK ADD  CONSTRAINT [Reftbl_user11] FOREIGN KEY([id_user])
REFERENCES [dbo].[tbl_user] ([id_user])
GO
ALTER TABLE [dbo].[tbl_userRol] CHECK CONSTRAINT [Reftbl_user11]
GO
USE [master]
GO
ALTER DATABASE [LMBA] SET  READ_WRITE 
GO
