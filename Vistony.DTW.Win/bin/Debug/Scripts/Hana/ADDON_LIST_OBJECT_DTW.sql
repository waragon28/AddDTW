CREATE PROCEDURE "ADDON_LIST_OBJECT_DTW" (in Tipo_Data nvarchar(50)) 
AS 
BEGIN /*ITEM 1*/ 	

IF :Tipo_Data='Datos de configuración' THEN
			  
	SELECT 'Datos definidos por el usuario' as "Objecto",
		   'Tabla (UDT)' as "Level_1_2",
		   'U_'||''||"TableName" AS "Level_1_2_3", "Code" as "ID_OBJ"
	FROM "OUDO"  ;
END IF;

IF :Tipo_Data='Master Data' THEN	 			
	SELECT 	DISTINCT 'Master Data' as "Objecto",
			'Objecto (UDO)' as "Level_1_2",
			'UDO_'||"Code" AS "Level_1_2_3", "Code" as "ID_OBJ"
	FROM "OUDO";
END IF;	
   
END;