CREATE PROCEDURE dbo.sp_InsertIntoUsuariosMultiple
(
     @ExcelData AS dbo.YourExcelTableType READONLY
)
AS
BEGIN
    SET NOCOUNT ON;

    -- Insertar la fila en la tabla "usuarios"
    INSERT INTO usuarios (nombre, apellido, correo, sexo)
     SELECT nombre, apellido, correo, sexo
    FROM @ExcelData;
END