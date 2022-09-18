# Universidad Rafael Landívar
# CHW, S.A.
# Procesamiento Cuadros de Oferta y Utilización
# Años 2013-2020
# Consultor Renato Vargas

# Llamar librerías
library(readxl)
library(openxlsx)
library(reshape2)
library(stringr)
library(plyr)
library(RSQLite)
library(DBI)

# Limpiar el área de trabajo
rm(list = ls())

# Se pone la ruta de trabajo en una variable (con "/")
wd <- "C:/Users/renato/GitHub/scn_scae_gt/datos/"
# Cambiar la ruta de trabajo con la variable anterior
setwd(wd)
getwd()

# Lógica recursiva
# ================

# Limpiar el área de trabajo
rm(list = ls())

archivo <- "GTM_COUS_DESAGREGADOS_2013_2020.xlsx"
hojas <- excel_sheets(archivo)

# Eliminamos las de precios constantes por conveniencia de análisis
hojas <- hojas[-c(3,5,7,9,11,13,15)]


lista <- c("inicio")

for (i in 1:length(hojas)) {
  # Extraemos el año y la unidad de medida
  info <- read_excel(
    archivo,
    range = paste("'" , hojas[i], "'!A8:B9", sep = "") ,
    col_names = FALSE,
    col_types = "text",
  )
  # Extraemos el texto de la cadena de caracteres
  anio <- as.numeric((str_extract(info[1,2 ], "\\d{4}")))
  unidad <- toString(info[2,1 ]) # unidad de medida
  # precios Corrientes o medidas Encadenados
  if (unidad != "Millones de quetzales") {
    precios <- "Constantes"
    id_precios <- 2
    unidad <- c("Millones de quetzales en medidas encadenadas de volumen con año de referencia 2013")
    id_unidad <- 2
  }
  else {
    precios <- "Corrientes"
    id_precios <- 1
    id_unidad <- 1
  }
  
  # Cuadro de Oferta
  # ================
  
  oferta <- as.matrix(read_excel(
    archivo,
    range = paste("'" , hojas[i], "'!C16:FL240", sep = ""),
    col_names = FALSE,
    col_types = "numeric"
  ))
  rownames(oferta) <- c(sprintf("f%03d", seq(1, dim(oferta)[1])))
  colnames(oferta) <- c(sprintf("oc%03d", seq(1, dim(oferta)[2])))
  
  # Columnas a eliminar con subtotales y totales
  
  # 129 SIFMI (Lo eliminamos porque no tiene datos en ningún año)
  # 130	P1 PRODUCCION (PB)	Producción de mercado		SUBTOTAL DE MERCADO
  # 135	P1 PRODUCCION (PB)	Producción para uso final propio		SUBTOTAL USO FINAL PROPIO
  # 147	P1 PRODUCCION (PB)	Producción no de mercado		SUBTOTAL NO DE MERCADO
  # 148	P1 TOTAL PRODUCCION (PB)
  # 149 VACIA
  # 150 VACIA
  # 153	P7 IMPORTACIONES (CIF)	TOTAL		TOTAL
  # 155 TOTAL OFERTA (a precios básicos)
  # 160	D21 IMPUESTOS SOBRE PRODUCTOS	TOTAL		TOTAL
  # 165 MARGENES DE DISTRIBUCION TOTAL
  # 166	TOTAL OFERTA (PC)	TOTAL OFERTA (PC)
  
  # Filas a eliminar con celdas vacías y subtotales
  
  # 214, 215, 216, 217, 218, 219, 224, 225
  
  oferta <- oferta[-c(214, 215, 216, 217, 218, 
                      219, 224, 225),
                   -c(129, 130,135,147,148,149,
                      150,153,155,160,165,166)]
  
  # Desdoblamos
  oferta <- cbind(anio, id_precios,1, melt(oferta), id_unidad)
  
  colnames(oferta) <-
    c("anio",
      "id_precios",
      "id_cuadro",
      "id_fila",
      "id_columna",
      "valor",
      "id_unidad")
  
  # Cuadro de utilización
  # =====================
  
  utilizacion <- as.matrix(read_excel(
    archivo,
    range = paste("'" , hojas[i], "'!C252:FL476", sep = ""),
    col_names = FALSE,
    col_types = "numeric"
  ))
  rownames(utilizacion) <-
    c(sprintf("f%03d", seq(1, dim(utilizacion)[1])))
  colnames(utilizacion) <-
    c(sprintf("uc%03d", seq(1, dim(utilizacion)[2])))
  
  #   Columnas y filas a eliminar con subtotales y totales
  
  #   uc130	P2 CONSUMO INTERMEDIO (PC)	SUBTOTAL DE MERCADO
  #   uc135	P2 CONSUMO INTERMEDIO (PC)	SUBTOTAL USO FINAL PROPIO
  #   uc147	P2 CONSUMO INTERMEDIO (PC)	SUBTOTAL NO DE MERCADO
  #   uc148	"P2 TOTAL CONSUMO INTERMEDIO
  #   uc149 VACIA
  #   uc150 VACIA
  #   uc153	P6 EXPORTACIONES (FOB) TOTAL
  #   uc157 SUBTOTAL ISFL
  #   uc160 Subtotal Gobienro General
  #   uc161 Total gasto de consumo final
  #   uc165 P.5 TOTAL FORMACION BRUTA DE CAPITAL
  #   uc166	TOTAL UTILIZACIÓN
  
  #   Filas a eliminar
  #   214, 215, 216, 218, 219, 224, 225
  
  utilizacion <- utilizacion[-c(214, 215, 216, 218, 219, 224, 225),
                             -c(129, 130, 135, 147, 148, 149, 150,
                                153, 157, 160, 161, 165, 166)]
  
  # Desdoblamos
  utilizacion <-
    cbind(anio, id_precios,2, melt(utilizacion), id_unidad)
  
  colnames(utilizacion) <-
    c("anio",
      "id_precios",
      "id_cuadro",
      "id_fila",
      "id_columna",
      "valor",
      "id_unidad")

  
  # Unimos todas las partes
  if (precios == "Corrientes") {
    union <- rbind(oferta, 
                   utilizacion
                   #,valorAgregado, 
                   #empleo
                   )
    
    assign(paste("COU_", anio, "_", precios, sep = ""), 
           union)
  } else {
    union <- rbind(oferta, utilizacion)
    assign(paste("COU_", anio, "_", precios, sep = ""), 
           union)
  }
  lista <- c(lista, paste("COU_", anio, "_", precios, sep = ""))
}
# Actualizamos nuestra lista de objetos creados
lista <- lapply(lista[-1], as.name)

# Unimos los objetos de todos los años y precios
SCN <- do.call(rbind.data.frame, lista)

# Convertimos los valores NA a 0
SCN$valor[is.na(SCN$valor)] <- 0

# Y borramos los objetos individuales
do.call(rm,lista)

# Recolección de basura
gc()

# ===============
# Clasificaciones
#================

clasifs <- "CLASIFICACIONES.xlsx"

# Columnas
columnas <- read_excel(
  clasifs,
  sheet = "columnas"  ,
  col_names = TRUE
)

# filas
filas <- read_excel(
  clasifs,
  sheet = "filas"  ,
  col_names = TRUE
)

# ntg2
ntg2 <- read_excel(
  clasifs,
  sheet = "ntg2"  ,
  col_names = TRUE
)

# npg
npg <- read_excel(
  clasifs,
  sheet = "npg"  ,
  col_names = TRUE
)

# naeg
naeg <- read_excel(
  clasifs,
  sheet = "naeg"  ,
  col_names = TRUE
)

# areas_filas
areas_filas <- read_excel(
  clasifs,
  sheet = "areas_filas"  ,
  col_names = TRUE
)

# areas_columnas
areas_columnas <- read_excel(
  clasifs,
  sheet = "areas_columnas"  ,
  col_names = TRUE
)

# areas_columnas
energia <- read_excel(
  clasifs,
  sheet = "energia"  ,
  col_names = TRUE
)

# cuadros
cuadros <- read_excel(
  clasifs,
  sheet = "cuadros"  ,
  col_names = TRUE
)


# ====================
# Base de datos SQLite
# ====================

if (file.exists("scn.db")) {
  file.remove("scn.db")
}
con <- dbConnect(RSQLite::SQLite(), "scn.db")
dbCreateTable(con, "oferta_utilizacion", SCN)
dbAppendTable(con, "oferta_utilizacion", SCN)
summary(con)

# columnas
dbCreateTable(con, "columnas", columnas)
dbAppendTable(con, "columnas", columnas)

# filas
dbCreateTable(con, "filas", filas)
dbAppendTable(con, "filas", filas)

# ntg2
dbCreateTable(con, "ntg2", ntg2)
dbAppendTable(con, "ntg2", ntg2)

# npg
dbCreateTable(con, "npg", npg)
dbAppendTable(con, "npg", npg)

# naeg
dbCreateTable(con, "naeg", naeg)
dbAppendTable(con, "naeg", naeg)

# areas_filas
dbCreateTable(con, "areas_filas", areas_filas)
dbAppendTable(con, "areas_filas", areas_filas)

# areas_columnas
dbCreateTable(con, "areas_columnas", areas_columnas)
dbAppendTable(con, "areas_columnas", areas_columnas)

# energia
dbCreateTable(con, "energia", energia)
dbAppendTable(con, "energia", energia)

# cuadros
dbCreateTable(con, "cuadros", cuadros)
dbAppendTable(con, "cuadros", cuadros)

# Consulta prueba
dbGetQuery(con, "
SELECT
    anio,
    cuadro,
    sum(valor)
FROM 
    oferta_utilizacion
JOIN
    cuadros
ON
    oferta_utilizacion.id_cuadro
    =
    cuadros.id_cuadro
GROUP BY 
    anio, oferta_utilizacion.id_cuadro
ORDER BY
    anio, oferta_utilizacion.id_cuadro
")

dbListTables(con)

dbDisconnect(con)

# ====================
# Otras exportaciones
# ====================

# Y lo exportamos a Excel
# write.xlsx(
#   SCN,
#   "SCN_BD.xlsx",
#   sheetName= "SCNGT_BD",
#   rowNames=FALSE,
#   colnames=FALSE,
#   overwrite = TRUE,
#   asTable = FALSE
# )

# El formato CSV se exporta muy grande, pero se comprime muy bien a 3mb
# write.csv(
#   SCN,
#   "scn.csv",
#   col.names = TRUE,
#   row.names = FALSE,
#   fileEncoding = "UTF-8"
# )
