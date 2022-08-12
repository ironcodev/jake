declare @messages table([key] varchar(100), [value] nvarchar(1000))

insert into @messages([key], [value]) values
('/en/account/signin/NoModel', N'Unexpected error. No signin model!'),
('/en/account/signin/NoUserName', N'Please enter your username.')

insert into dbo.Messages
select [key], [value] from @messages where [key] not in
(select [key] from dbo.Messages)
