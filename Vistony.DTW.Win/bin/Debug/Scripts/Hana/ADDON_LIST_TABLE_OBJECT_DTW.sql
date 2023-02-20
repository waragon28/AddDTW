CREATE PROCEDURE "ADDON_LIST_TABLE_OBJECT_DTW" (IN TIPO_OBJ NVARCHAR(50),in CodeObjDTW nvarchar(50)) 
AS 
BEGIN /*ITEM 1*/ 	
IF :TIPO_OBJ='Master Data' THEN
 
SELECT 	T0."TableName" as "Tabla",'' as "Archivo" FROM "OUDO" T0
				WHERE T0."Code"=:CodeObjDTW
				UNION ALL  
SELECT T1."TableName",'' as "Archivo"  FROM "UDO1" T1
				WHERE T1."Code"=:CodeObjDTW; 
END IF;

IF :TIPO_OBJ='Datos de configuración'  THEN 
SELECT :CodeObjDTW AS "Tabla",'' as "Archivo"  FROM "DUMMY" T1;
END IF;
END;