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
library(readr)

# Limpiar el área de trabajo
rm(list = ls())

# Lógica recursiva
# ================

# Nota: Hacemos referencia a los archivos tomando 
# como punto de partida la raíz del repositorio

# scn_scae_gt
#   ├───datos
#   ├───documentos
#   │   └───presentaciones
#   └───scripts


# Limpiar el área de trabajo
rm(list = ls())

# Identificador de país
pais <- "Guatemala"
iso3 <- "GTM"

archivo <- "datos/GTM_COUS_DESAGREGADOS_2013_2020.xlsx"
hojas <- excel_sheets(archivo)

# Eliminamos las de precios constantes por conveniencia de análisis
hojas <- hojas[-c(3,5,7,9,11,13,15)]
# i <- 1

lista <- c("inicio")

for (i in 1:length(hojas)) {
  
  # ==================================
  # MÓDULO DE PROCESAMIENTO DE EXCELES
  # ==================================
  
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
  
  # Determinación de precios corrientes o constantes
  if (unidad != "Millones de quetzales") {
    precios <- "Constantes"
    id_precios <- 2
    unidad <- c("Millones de quetzales en medidas encadenadas de volumen con año de referencia 2013")
    id_unidad <- 2
  }else {
    precios <- "Corrientes"
    id_precios <- 1
    id_unidad <- 1 # Millones de quetzales
  }
  
  # Procesamiento del Cuadro de Oferta Energética
  # ==================================
  
  id_cuadro <- 1 # 1  Oferta monetaria
  
  # Como puede verse en el parámetro range de la función read_excel(), 
  # lo que hacemos, en esencia, es leer todas las celdas numéricas del
  # cuadro, incluyendo espacios vacíos, columnas y filas de subtotales,
  # para poder clasificar cada celda respecto a cuatro dimensiones:
  # la transacción a la que corresponde, la actividad económica que la 
  # realiza, el producto objeto de la transacción y una cuarta que
  # captura los elementos no estándar comunes en los cuadros de oferta
  # y utilización publicados por los bancos centrales e institutos de
  # estadística a través de lo que llamamos "Áreas transaccionales de
  # fila o de columna". 
  
  # Leemos todo el rectángulo que contiene datos.
  oferta <- as.matrix(read_excel(
    archivo,
    range = paste("'" , hojas[i], "'!C16:FL240", sep = ""),
    col_names = FALSE,
    col_types = "numeric"
  ))
  
  # Y en este punto simplemente le damos una identificación correlativa 
  # a cada fila y a cada columna. En el caso de las filas, ambos cuadros 
  # de oferta y utilización, comparten exactamente los mismos valores. 
  # Por esa razón el correlativo solamente se compone de un identificador 
  # ISO3 de país (GTM en este caso), una letra "f" minúscula (de filas)  
  # y un correlativo de tres dígitos. Esto lo hacemos accediendo a los 
  # nombres de fila y columna con las funciones `rownames()` y `colnames()`
  # y asignándoles la secuencia compuesta por los elementos descritos.
  
  # Las funciones esenciales para lograr esto son `sprintf()` combinada con
  # 'seq()'. Nótese que se embebe la función `paste()` que nos permite
  # concatenar el texto del código de país (GTM por ejemplo) con el formato
  # del correlativo de fila o columna. También embebemos la función `dim()`
  # dentro de la función `seq()` la cual nos permite saber programáticamente
  # las dimensiones de la matriz importada.
  
  rownames(oferta) <- c(sprintf(paste(iso3,"f%03d",  sep = ""), seq(1, dim(oferta)[1])))
  colnames(oferta) <- c(sprintf(paste(iso3,"oc%03d", sep = ""), seq(1, dim(oferta)[2])))

  # El siguiente punto consiste en identificar las columnas a eliminar,
  # las cuales contienen subtotales y totales, contando como 1, la primera
  # celda leída en el módulo anterior.
  
  # 129 SIFMI (Lo eliminamos porque no tiene datos en ningún año)
  # 130	P1 PRODUCCION (PB)	Producción de mercado	SUBTOTAL DE MERCADO
  # 135	P1 PRODUCCION (PB)	Producción para uso final propio SUBTOTAL USO FINAL PROPIO
  # 147	P1 PRODUCCION (PB)	Producción no de mercado SUBTOTAL NO DE MERCADO
  # 148	P1 TOTAL PRODUCCION (PB)
  # 149 VACIA
  # 150 VACIA
  # 153	P7 IMPORTACIONES (CIF)	TOTAL		TOTAL
  # 155 TOTAL OFERTA (a precios básicos)
  # 160	D21 IMPUESTOS SOBRE PRODUCTOS	TOTAL		TOTAL
  # 165 MARGENES DE DISTRIBUCION TOTAL
  # 166	TOTAL OFERTA (PC)	TOTAL OFERTA (PC)
  
  # Seguidamnete, identificamos en nuestro Excel las filas vacías y que contienen
  # subtotales para ser eliminadas.
  
  # 214, 215, 216, 218, 219, 224, 225
  
  # A través del concepto de índices [fila, columna], utilizado en la notación 
  # matricial para referirnos a la posición de una o más celdas en una matriz, 
  # a través del signo menos (-) eliminamos las filas y columnas identificadas 
  # en el paso anterior.
  
  oferta <- oferta[-c(214, 215, 216, 218, 
                      219, 224, 225),
                   -c(129, 130,135,147,148,149,
                      150,153,155,160,165,166)]
  
  # El resultado es una matriz rectangular de datos válidos y únicos. Es decir,
  # sin subtotales y totales que duplican los datos básicos. El siguiente paso 
  # transforma esa matriz de dos dimensiones a una tabla de datos de una sola
  # dimensión que nos va a permitir guardar y manipular nuestros datos a 
  # través de una base de datos y el lenguaje de consultas estructuradas SQL.
  
  # Esto se logra con la función `melt()` de la librería `reshape2`. Nótese
  # que para ahorrar un paso creamos columnas adicionales que representan
  # lo mínimo necesario para poder identificar el año, el identificador de
  # a qué cuadro pertenecen los datos y la unidad de medida del cuadro.
  
  # Anteriormente, agregábamos un identificador de precios constantes o 
  # corrientes. No obstante, en aras de estandarizar este procedimiento
  # para temas energéticos, ambientales de empleo y otros que no usan esa 
  # diferenciación, decidimos que la distinción de constantes o corrientes 
  # se hará a través del campo de unidad de medida `id_unidad`. Es decir
  # Millones de quetzales constantes. 
  
  # Por conveniencia no se incluirá precios constantes.
  
  oferta <- cbind(iso3, anio,id_cuadro, melt(oferta), id_unidad)
  
  colnames(oferta) <-
    c("iso3",
      "anio",
      "id_cuadro",
      "id_fila",
      "id_columna",
      "valor",
      "id_unidad")
  
  
  # Procesamiento del Cuadro de Utilización
  # =======================================
  
  id_cuadro  <- 2 # Utilización monetaria
  
  # El procesamiento del Cuadro de Utilización es idéntico al de oferta. 
  
  # Leemos el rectángulo de datos del archivo de Excel original.
  utilizacion <- as.matrix(read_excel(
    archivo,
    range = paste("'" , hojas[i], "'!C252:FL476", sep = ""),
    col_names = FALSE,
    col_types = "numeric"
  ))
  
  # Y nombramos las filas y columnas con los correlativos necesarios.
  rownames(utilizacion) <-
    c(sprintf(paste(iso3,"f%03d",  sep = ""), seq(1, dim(utilizacion)[1])))
  
  colnames(utilizacion) <-
    c(sprintf(paste(iso3,"uc%03d", sep = ""), seq(1, dim(utilizacion)[2])))
  
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
  
  # A través de índices eliminamos lo que no se necesita.
  utilizacion <- utilizacion[-c(214, 215, 216, 218, 219, 224, 225),
                             -c(129, 130, 135, 147, 148, 149, 150,
                                153, 157, 160, 161, 165, 166)]
  
  # Aplicamos la función `melt()` para convertir a formato de tabla de base
  # de datos, poniendo cuidado en que las columnas sean compatibles con el
  # cuadro de oferta que creamos en la sección anterior.
  
  utilizacion <-
    cbind(iso3,anio, id_cuadro, melt(utilizacion), id_unidad)
  
  colnames(utilizacion) <-
    c("iso3",
      "anio",
      "id_cuadro",
      "id_fila",
      "id_columna",
      "valor",
      "id_unidad")

  
  # Procesamiento del Valor Agregado
  # =======================================
  
  id_cuadro  <- 10 # Valor Agregado
  
  # El procesamiento del Cuadro de Utilización es idéntico al de oferta. 
  
  # Leemos el rectángulo de datos del archivo de Excel original.
  valorAgregado <- as.data.frame(read_excel(
    archivo,
    range = paste("'" , hojas[i], "'!D478:ET478", sep = ""),
    col_names = FALSE,
    col_types = "numeric"
  ))
  
  # Y nombramos las filas y columnas con los correlativos necesarios.
  rownames(valorAgregado) <-
    "GTMva000"
  
  colnames(valorAgregado) <-
    c(sprintf(paste(iso3,"uc%03d", sep = ""), seq(1, dim(valorAgregado)[2])))
  
  #   Columnas y filas a eliminar con subtotales y totales
  
  #   uc130	P2 CONSUMO INTERMEDIO (PC)	SUBTOTAL DE MERCADO
  #   uc135	P2 CONSUMO INTERMEDIO (PC)	SUBTOTAL USO FINAL PROPIO
  #   uc147	P2 CONSUMO INTERMEDIO (PC)	SUBTOTAL NO DE MERCADO
  
  # A través de índices eliminamos lo que no se necesita.
  valorAgregado <- valorAgregado[ , -c(129, 130, 135, 147)]
  
  # Aplicamos la función `melt()` para convertir a formato de tabla de base
  # de datos, poniendo cuidado en que las columnas sean compatibles con el
  # cuadro de oferta que creamos en la sección anterior.
  
  valorAgregado <-
    cbind(iso3,anio, id_cuadro, "GTMva000", melt(valorAgregado), id_unidad)
  
  colnames(valorAgregado) <-
    c("iso3",
      "anio",
      "id_cuadro",
      "id_fila",
      "id_columna",
      "valor",
      "id_unidad")
  
  # Unión de los cuadros procesados
  # ===============================
  
  if (precios == "Corrientes") {
    union <- rbind(oferta, 
                   utilizacion,
                   valorAgregado, 
                   #empleo
                   deparse.level = 0)
    
    assign(paste("COU_", anio, "_", precios, sep = ""), 
           union)
  } else {
    union <- rbind(oferta, utilizacion, valorAgregado, deparse.level = 0)
    assign(paste("COU_", anio, "_", precios, sep = ""), 
           union)
  }
  lista <- c(lista, paste("COU_", anio, "_", precios, sep = ""))
}

# =====================================================
# MÓDULO DE CREACIÓN DE BASE DE DATOS RELACIONAL SQLITE
#======================================================

# Todas las funciones de procesamiento de datos del módulo anterior
# están embebidas en un bucle que itera a través de todas las pestañas
# de un archivo de Excel que contienen los cuadros de oferta y 
# utilización para cada año, unas a precios corrientes y otras a 
# precios constantes, y permite guardar cada grupo de cuadros procesados
# con un nombre según su año (por ejemplo, COU_2013_corrientes).

# Llevamos el control de los objetos a través de una `lista` de objetos
# que hemos ido actualizando a medida que se agrega cada nuevo objeto.
# Ahora esa lista nos servirá para hacer referencia a todos los objetos
# que queremos unir en uno solo.

# Eliminamos el primer elemento provisional que usamos para inicializar
# la lista, pues R no nos permite declarar listas vacías. 
lista <- lapply(lista[-1], as.name)

# Unimos los objetos de todos los años y precios
SCN <- do.call(rbind.data.frame, lista)

# Por conveniencia, convertimos los valores NA a 0
SCN$valor[is.na(SCN$valor)] <- 0

# Y borramos los objetos individuales
do.call(rm,lista)

# Recolección de basura
gc()

# ===============
# Clasificaciones
#================

# Ya le hemos dado identificadores únicos a cada una de las filas y columnas
# de nuestra información importada. Ahora es necesario darle sentido a cada
# una de ellas. Esto se hace a través de la creación de tablas de 
# equivalencias. A través del proceso de normalización de base de datos nos
# aseguramos que no existan redundancias en las tablas de la base de datos.

# La normalización permite mantener las descripciones de cada elemento en una
# tabla dedicada, para que estas solo existan en un lugar, de manera que si 
# se necesita hacer cambios, solamente se harán una vez.

# Hemos creado un archivo de excel que cuenta con estas tablas de equivalencia
# en cada una de sus pestañas. Aquí las cargamos.

# Nótese que aquí estamos creando una base de datos relacional y por eso 
# mantendremos separadas las tablas. Si estuviéramos haciendo un "Flat File"
# lo haríamos creando las columnas necesarias en la misma tabla a través de
# la función `join()`.

clasifs <- "datos/CLASIFICACIONES.xlsx"

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

# filas Energía
filasE <- read_excel(
  clasifs,
  sheet = "filasE"  ,
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

if (file.exists("salidas/scn.db")) {
  file.remove("salidas/scn.db")
}
con <- dbConnect(RSQLite::SQLite(), "salidas/scn.db")
dbCreateTable(con, "scn", SCN)
dbAppendTable(con, "scn", SCN)
summary(con)

# columnas
dbCreateTable(con, "columnas", columnas)
dbAppendTable(con, "columnas", columnas)

# filas
dbCreateTable(con, "filas", filas)
dbAppendTable(con, "filas", filas)

# filasE
dbCreateTable(con, "filasE", filasE)
dbAppendTable(con, "filasE", filasE)

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


dbListTables(con)
dbListFields(con,"cuadros")

# Creamos una vista para facilitar las consultas y exportación a flat file
# de Excel. Usamos un archivo SQL para trabajar más fácilmente con
# la consulta SQL aparte.

query <- read_file("scripts/VIEW_oferta_utilizacion.sql")
dbSendQuery(con, query)


# ====================
# Otras exportaciones
# ====================


SCN2 <- dbGetQuery(con,"
SELECT 
  * 
FROM 
  oferta_utilizacion 
")

# Y lo exportamos a Excel
write.xlsx(
  SCN2,
  "salidas/SCN_BD.xlsx",
  sheetName= "SCNGT_BD",
  rowNames=FALSE,
  colnames=FALSE,
  overwrite = TRUE,
  asTable = FALSE
)

dbDisconnect(con)

# El formato CSV se exporta muy grande, pero se comprime muy bien a 3mb
# write.csv(
#   SCN,
#   "scn.csv",
#   col.names = TRUE,
#   row.names = FALSE,
#   fileEncoding = "UTF-8"
# )
