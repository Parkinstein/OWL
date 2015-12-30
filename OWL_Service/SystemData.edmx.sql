
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, 2012 and Azure
-- --------------------------------------------------
-- Date Created: 12/30/2015 13:08:39
-- Generated from EDMX file: C:\Users\boris\Downloads\ASPNetMVCExtendingIdentity2Roles\OWL_Service\SystemData.edmx
-- --------------------------------------------------

SET QUOTED_IDENTIFIER OFF;
GO
USE [aspnetdb];
GO
IF SCHEMA_ID(N'dbo') IS NULL EXECUTE(N'CREATE SCHEMA [dbo]');
GO

-- --------------------------------------------------
-- Dropping existing FOREIGN KEY constraints
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[FK_dbo_AspNetUserClaims_dbo_AspNetUsers_UserId]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[AspNetUserClaims] DROP CONSTRAINT [FK_dbo_AspNetUserClaims_dbo_AspNetUsers_UserId];
GO
IF OBJECT_ID(N'[dbo].[FK_dbo_AspNetUserLogins_dbo_AspNetUsers_UserId]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[AspNetUserLogins] DROP CONSTRAINT [FK_dbo_AspNetUserLogins_dbo_AspNetUsers_UserId];
GO
IF OBJECT_ID(N'[dbo].[FK_dbo_AspNetUserRoles_dbo_AspNetRoles_Role_Id]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[AspNetUserRoles] DROP CONSTRAINT [FK_dbo_AspNetUserRoles_dbo_AspNetRoles_Role_Id];
GO
IF OBJECT_ID(N'[dbo].[FK_dbo_AspNetUserRoles_dbo_AspNetRoles_RoleId]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[AspNetUserRoles] DROP CONSTRAINT [FK_dbo_AspNetUserRoles_dbo_AspNetRoles_RoleId];
GO
IF OBJECT_ID(N'[dbo].[FK_dbo_AspNetUserRoles_dbo_AspNetUsers_ApplicationUser_Id]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[AspNetUserRoles] DROP CONSTRAINT [FK_dbo_AspNetUserRoles_dbo_AspNetUsers_ApplicationUser_Id];
GO
IF OBJECT_ID(N'[dbo].[FK_dbo_AspNetUserRoles_dbo_AspNetUsers_UserId]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[AspNetUserRoles] DROP CONSTRAINT [FK_dbo_AspNetUserRoles_dbo_AspNetUsers_UserId];
GO
IF OBJECT_ID(N'[dbo].[FK_MeetingAtendee_Meeting]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[MeetingAttendees] DROP CONSTRAINT [FK_MeetingAtendee_Meeting];
GO
IF OBJECT_ID(N'[dbo].[FK_Meetings_Meetings]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Meetings] DROP CONSTRAINT [FK_Meetings_Meetings];
GO

-- --------------------------------------------------
-- Dropping existing tables
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[AllVmrs]', 'U') IS NOT NULL
    DROP TABLE [dbo].[AllVmrs];
GO
IF OBJECT_ID(N'[dbo].[AspNetRoles]', 'U') IS NOT NULL
    DROP TABLE [dbo].[AspNetRoles];
GO
IF OBJECT_ID(N'[dbo].[AspNetUserClaims]', 'U') IS NOT NULL
    DROP TABLE [dbo].[AspNetUserClaims];
GO
IF OBJECT_ID(N'[dbo].[AspNetUserLogins]', 'U') IS NOT NULL
    DROP TABLE [dbo].[AspNetUserLogins];
GO
IF OBJECT_ID(N'[dbo].[AspNetUserRoles]', 'U') IS NOT NULL
    DROP TABLE [dbo].[AspNetUserRoles];
GO
IF OBJECT_ID(N'[dbo].[AspNetUsers]', 'U') IS NOT NULL
    DROP TABLE [dbo].[AspNetUsers];
GO
IF OBJECT_ID(N'[dbo].[Ivr_Themes]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Ivr_Themes];
GO
IF OBJECT_ID(N'[dbo].[MeetingAttendees]', 'U') IS NOT NULL
    DROP TABLE [dbo].[MeetingAttendees];
GO
IF OBJECT_ID(N'[dbo].[Meetings]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Meetings];
GO
IF OBJECT_ID(N'[dbo].[PrivatePhBs]', 'U') IS NOT NULL
    DROP TABLE [dbo].[PrivatePhBs];
GO
IF OBJECT_ID(N'[dbo].[Settings]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Settings];
GO
IF OBJECT_ID(N'[dbo].[VmrAliases]', 'U') IS NOT NULL
    DROP TABLE [dbo].[VmrAliases];
GO

-- --------------------------------------------------
-- Creating all tables
-- --------------------------------------------------

-- Creating table 'AllVmrs'
CREATE TABLE [dbo].[AllVmrs] (
    [Id] int  NOT NULL,
    [allow_guests] bit  NULL,
    [description] nvarchar(max)  NULL,
    [force_presenter_into_main] bit  NULL,
    [guest_pin] nvarchar(50)  NULL,
    [guest_view] nvarchar(max)  NULL,
    [host_view] nvarchar(max)  NULL,
    [max_callrate_in] nvarchar(50)  NULL,
    [max_callrate_out] nvarchar(50)  NULL,
    [name] nvarchar(max)  NULL,
    [participant_limit] nvarchar(50)  NULL,
    [pin] nvarchar(10)  NULL,
    [resource_uri] nvarchar(max)  NULL,
    [service_type] nvarchar(255)  NULL,
    [tag] nvarchar(max)  NULL,
    [vmid] int  NULL
);
GO

-- Creating table 'AspNetRoles'
CREATE TABLE [dbo].[AspNetRoles] (
    [Id] nvarchar(128)  NOT NULL,
    [Name] nvarchar(256)  NOT NULL,
    [Description] nvarchar(max)  NULL,
    [Discriminator] nvarchar(128)  NOT NULL
);
GO

-- Creating table 'AspNetUserClaims'
CREATE TABLE [dbo].[AspNetUserClaims] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [UserId] nvarchar(128)  NOT NULL,
    [ClaimType] nvarchar(max)  NULL,
    [ClaimValue] nvarchar(max)  NULL
);
GO

-- Creating table 'AspNetUserLogins'
CREATE TABLE [dbo].[AspNetUserLogins] (
    [LoginProvider] nvarchar(128)  NOT NULL,
    [ProviderKey] nvarchar(128)  NOT NULL,
    [UserId] nvarchar(128)  NOT NULL
);
GO

-- Creating table 'AspNetUserRoles'
CREATE TABLE [dbo].[AspNetUserRoles] (
    [UserId] nvarchar(128)  NOT NULL,
    [RoleId] nvarchar(128)  NOT NULL,
    [Discriminator] nvarchar(128)  NOT NULL,
    [Role_Id] nvarchar(128)  NULL,
    [ApplicationUser_Id] nvarchar(128)  NULL
);
GO

-- Creating table 'AspNetUsers'
CREATE TABLE [dbo].[AspNetUsers] (
    [Id] nvarchar(128)  NOT NULL,
    [Email] nvarchar(256)  NULL,
    [EmailConfirmed] bit  NOT NULL,
    [PasswordHash] nvarchar(max)  NULL,
    [SecurityStamp] nvarchar(max)  NULL,
    [PhoneNumber] nvarchar(max)  NULL,
    [PhoneNumberConfirmed] bit  NOT NULL,
    [TwoFactorEnabled] bit  NOT NULL,
    [LockoutEndDateUtc] datetime  NULL,
    [LockoutEnabled] bit  NOT NULL,
    [AccessFailedCount] int  NOT NULL,
    [UserName] nvarchar(256)  NOT NULL,
    [Sammaccount] nvarchar(max)  NULL,
    [Name] nvarchar(max)  NULL,
    [Surname] nvarchar(max)  NULL,
    [Position] nvarchar(max)  NULL,
    [Tel_int] nvarchar(max)  NULL,
    [Tel_ext] nvarchar(max)  NULL,
    [Tel_mob] nvarchar(max)  NULL,
    [DispName] nvarchar(max)  NULL,
    [Timezone] nvarchar(max)  NULL,
    [Sip_addr] nvarchar(max)  NULL,
    [H323_addr] nvarchar(max)  NULL,
    [Group] nvarchar(max)  NULL
);
GO

-- Creating table 'Ivr_Themes'
CREATE TABLE [dbo].[Ivr_Themes] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [name] nvarchar(max)  NULL,
    [uuid] nvarchar(max)  NULL,
    [vmid] int  NULL,
    [intid] int  NULL
);
GO

-- Creating table 'MeetingAttendees'
CREATE TABLE [dbo].[MeetingAttendees] (
    [MeetingID] int  NOT NULL,
    [AttendeeID] nvarchar(50)  NOT NULL
);
GO

-- Creating table 'Meetings'
CREATE TABLE [dbo].[Meetings] (
    [MeetingID] int IDENTITY(1,1) NOT NULL,
    [Start] datetime  NOT NULL,
    [End] datetime  NOT NULL,
    [Title] nvarchar(max)  NOT NULL,
    [Description] nvarchar(max)  NULL,
    [RoomID] int  NULL,
    [IsAllDay] bit  NOT NULL,
    [RecurrenceRule] nvarchar(max)  NULL,
    [RecurrenceID] int  NULL,
    [RecurrenceException] nvarchar(max)  NULL,
    [StartTimezone] nvarchar(max)  NULL,
    [EndTimezone] nvarchar(max)  NULL,
    [Oplink] nvarchar(max)  NULL,
    [AddAttend] nvarchar(max)  NULL,
    [FileLink] nvarchar(max)  NULL,
    [Record] bit  NOT NULL,
    [Recfile] nvarchar(max)  NULL,
    [InitName] nvarchar(max)  NULL,
    [InitFullname] nvarchar(max)  NULL,
    [reminder] bit  NOT NULL
);
GO

-- Creating table 'PrivatePhBs'
CREATE TABLE [dbo].[PrivatePhBs] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [OwSAN] nvarchar(max)  NULL,
    [IdREC] nvarchar(max)  NULL,
    [Group] nvarchar(50)  NULL
);
GO

-- Creating table 'Settings'
CREATE TABLE [dbo].[Settings] (
    [Id] int  NOT NULL,
    [AuthDnAddress] nvarchar(max)  NULL,
    [OU] nvarchar(max)  NULL,
    [UserGroup] nvarchar(max)  NULL,
    [AdminGroup] nvarchar(max)  NULL,
    [DnAdminUn] nvarchar(max)  NULL,
    [DnAdminPass] nvarchar(max)  NULL,
    [CobaMngAddress] nvarchar(max)  NULL,
    [CobaCfgAddress] nvarchar(max)  NULL,
    [CobaRecordsAddress] nvarchar(max)  NULL,
    [CobaRecLogin] nvarchar(max)  NULL,
    [CobaRecPass] nvarchar(max)  NULL,
    [CobaRecBdName] nvarchar(max)  NULL,
    [CobaRecBdTable] nvarchar(max)  NULL,
    [SmtpServer] nvarchar(max)  NULL,
    [SmtpPort] int  NULL,
    [SmtpSSL] bit  NOT NULL,
    [SmtpLogin] nvarchar(max)  NULL,
    [SmtpPassword] nvarchar(max)  NULL,
    [MailFrom_email] nvarchar(max)  NULL,
    [MailFrom_name] nvarchar(max)  NULL,
    [CobaMngLogin] nvarchar(max)  NULL,
    [CobaMngPass] nvarchar(max)  NULL
);
GO

-- Creating table 'VmrAliases'
CREATE TABLE [dbo].[VmrAliases] (
    [Id] int  NOT NULL,
    [alias] nvarchar(max)  NULL,
    [conference] nvarchar(255)  NULL,
    [description] nvarchar(max)  NULL,
    [vmid] int  NULL
);
GO

-- --------------------------------------------------
-- Creating all PRIMARY KEY constraints
-- --------------------------------------------------

-- Creating primary key on [Id] in table 'AllVmrs'
ALTER TABLE [dbo].[AllVmrs]
ADD CONSTRAINT [PK_AllVmrs]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'AspNetRoles'
ALTER TABLE [dbo].[AspNetRoles]
ADD CONSTRAINT [PK_AspNetRoles]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'AspNetUserClaims'
ALTER TABLE [dbo].[AspNetUserClaims]
ADD CONSTRAINT [PK_AspNetUserClaims]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [LoginProvider], [ProviderKey], [UserId] in table 'AspNetUserLogins'
ALTER TABLE [dbo].[AspNetUserLogins]
ADD CONSTRAINT [PK_AspNetUserLogins]
    PRIMARY KEY CLUSTERED ([LoginProvider], [ProviderKey], [UserId] ASC);
GO

-- Creating primary key on [UserId], [RoleId] in table 'AspNetUserRoles'
ALTER TABLE [dbo].[AspNetUserRoles]
ADD CONSTRAINT [PK_AspNetUserRoles]
    PRIMARY KEY CLUSTERED ([UserId], [RoleId] ASC);
GO

-- Creating primary key on [Id] in table 'AspNetUsers'
ALTER TABLE [dbo].[AspNetUsers]
ADD CONSTRAINT [PK_AspNetUsers]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Ivr_Themes'
ALTER TABLE [dbo].[Ivr_Themes]
ADD CONSTRAINT [PK_Ivr_Themes]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [MeetingID], [AttendeeID] in table 'MeetingAttendees'
ALTER TABLE [dbo].[MeetingAttendees]
ADD CONSTRAINT [PK_MeetingAttendees]
    PRIMARY KEY CLUSTERED ([MeetingID], [AttendeeID] ASC);
GO

-- Creating primary key on [MeetingID] in table 'Meetings'
ALTER TABLE [dbo].[Meetings]
ADD CONSTRAINT [PK_Meetings]
    PRIMARY KEY CLUSTERED ([MeetingID] ASC);
GO

-- Creating primary key on [Id] in table 'PrivatePhBs'
ALTER TABLE [dbo].[PrivatePhBs]
ADD CONSTRAINT [PK_PrivatePhBs]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Settings'
ALTER TABLE [dbo].[Settings]
ADD CONSTRAINT [PK_Settings]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'VmrAliases'
ALTER TABLE [dbo].[VmrAliases]
ADD CONSTRAINT [PK_VmrAliases]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- Creating foreign key on [Role_Id] in table 'AspNetUserRoles'
ALTER TABLE [dbo].[AspNetUserRoles]
ADD CONSTRAINT [FK_dbo_AspNetUserRoles_dbo_AspNetRoles_Role_Id]
    FOREIGN KEY ([Role_Id])
    REFERENCES [dbo].[AspNetRoles]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_dbo_AspNetUserRoles_dbo_AspNetRoles_Role_Id'
CREATE INDEX [IX_FK_dbo_AspNetUserRoles_dbo_AspNetRoles_Role_Id]
ON [dbo].[AspNetUserRoles]
    ([Role_Id]);
GO

-- Creating foreign key on [RoleId] in table 'AspNetUserRoles'
ALTER TABLE [dbo].[AspNetUserRoles]
ADD CONSTRAINT [FK_dbo_AspNetUserRoles_dbo_AspNetRoles_RoleId]
    FOREIGN KEY ([RoleId])
    REFERENCES [dbo].[AspNetRoles]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_dbo_AspNetUserRoles_dbo_AspNetRoles_RoleId'
CREATE INDEX [IX_FK_dbo_AspNetUserRoles_dbo_AspNetRoles_RoleId]
ON [dbo].[AspNetUserRoles]
    ([RoleId]);
GO

-- Creating foreign key on [UserId] in table 'AspNetUserClaims'
ALTER TABLE [dbo].[AspNetUserClaims]
ADD CONSTRAINT [FK_dbo_AspNetUserClaims_dbo_AspNetUsers_UserId]
    FOREIGN KEY ([UserId])
    REFERENCES [dbo].[AspNetUsers]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_dbo_AspNetUserClaims_dbo_AspNetUsers_UserId'
CREATE INDEX [IX_FK_dbo_AspNetUserClaims_dbo_AspNetUsers_UserId]
ON [dbo].[AspNetUserClaims]
    ([UserId]);
GO

-- Creating foreign key on [UserId] in table 'AspNetUserLogins'
ALTER TABLE [dbo].[AspNetUserLogins]
ADD CONSTRAINT [FK_dbo_AspNetUserLogins_dbo_AspNetUsers_UserId]
    FOREIGN KEY ([UserId])
    REFERENCES [dbo].[AspNetUsers]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_dbo_AspNetUserLogins_dbo_AspNetUsers_UserId'
CREATE INDEX [IX_FK_dbo_AspNetUserLogins_dbo_AspNetUsers_UserId]
ON [dbo].[AspNetUserLogins]
    ([UserId]);
GO

-- Creating foreign key on [ApplicationUser_Id] in table 'AspNetUserRoles'
ALTER TABLE [dbo].[AspNetUserRoles]
ADD CONSTRAINT [FK_dbo_AspNetUserRoles_dbo_AspNetUsers_ApplicationUser_Id]
    FOREIGN KEY ([ApplicationUser_Id])
    REFERENCES [dbo].[AspNetUsers]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_dbo_AspNetUserRoles_dbo_AspNetUsers_ApplicationUser_Id'
CREATE INDEX [IX_FK_dbo_AspNetUserRoles_dbo_AspNetUsers_ApplicationUser_Id]
ON [dbo].[AspNetUserRoles]
    ([ApplicationUser_Id]);
GO

-- Creating foreign key on [UserId] in table 'AspNetUserRoles'
ALTER TABLE [dbo].[AspNetUserRoles]
ADD CONSTRAINT [FK_dbo_AspNetUserRoles_dbo_AspNetUsers_UserId]
    FOREIGN KEY ([UserId])
    REFERENCES [dbo].[AspNetUsers]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating foreign key on [MeetingID] in table 'MeetingAttendees'
ALTER TABLE [dbo].[MeetingAttendees]
ADD CONSTRAINT [FK_MeetingAtendee_Meeting]
    FOREIGN KEY ([MeetingID])
    REFERENCES [dbo].[Meetings]
        ([MeetingID])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating foreign key on [RecurrenceID] in table 'Meetings'
ALTER TABLE [dbo].[Meetings]
ADD CONSTRAINT [FK_Meetings_Meetings]
    FOREIGN KEY ([RecurrenceID])
    REFERENCES [dbo].[Meetings]
        ([MeetingID])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_Meetings_Meetings'
CREATE INDEX [IX_FK_Meetings_Meetings]
ON [dbo].[Meetings]
    ([RecurrenceID]);
GO

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------