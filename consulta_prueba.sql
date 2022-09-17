SELECT 
    id_cuadro, 
    id_ntg2, 
    sum(valor)
FROM 
    oferta_utilizacion
LEFT JOIN 
    columnas
ON
    oferta_utilizacion.id_columna = columnas.id_columna 
LEFT JOIN 
    ntg2 
ON
    id_ntg2=idNTG2
WHERE 
    anio=2014
GROUP BY 
    id_cuadro, id_ntg2, ntg2
    
    
SELECT
    anio,
    id_cuadro, 
    sum(valor)
FROM 
    oferta_utilizacion
JOIN 
    columnas
GROUP BY 
    anio, id_ntg2, ntg2
