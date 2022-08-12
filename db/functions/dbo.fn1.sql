create or alter function dbo.fn1(@a int)
returns int
as
begin
    return @a * @a
end
go
