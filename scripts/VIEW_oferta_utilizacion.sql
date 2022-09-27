CREATE VIEW oferta_utilizacion AS
SELECT
    scn.iso3,
    scn.anio,
    scn.id_cuadro,
	cuadros.cuadro,
	ntg2.corr_ntg2,
	columnas.id_ntg2,
	ntg2.ntg2,
	ntg2.corr_ntg,
	ntg2.id_ntg,
	ntg2.ntg,
	ntg2.id_area_columnas,
	areas_columnas.area_columnas,
	areas_columnas.id_area_columnas_compactas,
	areas_columnas.area_columnas_compactas,
	columnas.id_naeg,
	naeg.naeg,
	filas.id_area_filas,
	areas_filas.area_filas,
	filas.id_npg4,
	npg.npg4,
	npg.id_energia,
	energia.energia,
	scn.valor
FROM 
    scn
LEFT JOIN 
    columnas ON
    scn.id_columna = columnas.id_columna
LEFT JOIN
    filas ON
    scn.id_fila = filas.id_fila
LEFT JOIN
    ntg2 ON
    columnas.id_ntg2 = ntg2.id_ntg2
LEFT JOIN
    npg ON
    filas.id_npg4 = npg.id_npg4
LEFT JOIN
	naeg ON	
	columnas.id_naeg = naeg.id_naeg
LEFT JOIN
	cuadros ON
	scn.id_cuadro = cuadros.id_cuadro
LEFT JOIN
	areas_filas ON
	filas.id_area_filas = areas_filas.id_area_filas
LEFT JOIN
	areas_columnas ON
	ntg2.id_area_columnas = areas_columnas.id_area_columnas
LEFT JOIN
	energia ON
	npg.id_energia = energia.id_energia