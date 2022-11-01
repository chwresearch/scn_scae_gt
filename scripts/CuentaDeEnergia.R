# Universidad Rafael Landívar
# CHW, S.A.
# Cuenta De Energía
# Años 2013-2020
# Consultor Renato Vargas

# Llamar librerías
library(readxl)
library(openxlsx)
library(reshape2)
library(readr)
library(plyr)
library(stringr)
library(RSQLite)
library(DBI)
library(pivottabler)

# Limpiar el área de trabajo
rm(list = ls())

# Lógica recursiva
# ================

# Nota: Hacemos referencia a los archivos tomando 
# como punto de partida la raíz del repositorio
# Para que esto funcione, es necesario abrir el 
# proyecto a través del proyecto de RStudio
# scn_scae_gt.Rproj

# scn_scae_gt
#   ├───datos
#   ├───documentos
#   │   └───presentaciones
#   └───scripts

# Identificador de país
pais <- "Guatemala"
iso3 <- "GTM"

archivo <- "datos/GTM_COUS_DESAGREGADOS_2013_2020.xlsx"
hojas <- excel_sheets(archivo)

# Eliminamos las de precios constantes por conveniencia de análisis
hojas <- hojas[-c(3,5,7,9,11,13,15)]
i <- 1

lista <- c("inicio")

# Antes de entrar al bucle de cálculo, importamos los 
# Coeficientes de emisiones, para no hacerlo cada
# ciclo.

# Coeficientes de CO2

CO2_coef <- as.matrix(read_excel(
  "datos/CLASIFICACIONES.xlsx",
  range = "gei'!C6:EY21",
  col_names = FALSE,
  col_types = "numeric"
))

CH4_coef <- as.matrix(read_excel(
  "datos/CLASIFICACIONES.xlsx",
  range = "gei'!C25:EY40",
  col_names = FALSE,
  col_types = "numeric"
))

N2O_coef <- as.matrix(read_excel(
  "datos/CLASIFICACIONES.xlsx",
  range = "gei'!C44:EY59",
  col_names = FALSE,
  col_types = "numeric"
))

# Abrimos la conexión a la base de datos
con <- dbConnect(RSQLite::SQLite(), "salidas/scn.db")


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
  
  id_cuadro <- 3 # 1  Oferta energética monetaria
  
  # Leemos todo el rectángulo que contiene datos.
  oferta <- as.matrix(read_excel(
    archivo,
    range = paste("'" , hojas[i], "'!C16:FL240", sep = ""),
    col_names = FALSE,
    col_types = "numeric"
  ))
  
  # Y en este punto simplemente le damos una identificación correlativa 
  # a cada fila y a cada columna. 
  
  rownames(oferta) <- c(sprintf(paste(iso3,"f%03d",  sep = ""), seq(1, dim(oferta)[1])))
  colnames(oferta) <- c(sprintf(paste(iso3,"oc%03d", sep = ""), seq(1, dim(oferta)[2])))
  
  # A través de índices extraemos solo las filas que nos interesan
  # (productos energéticos) y eliminamos las columnas que no necesitamos
  # con totales y subtotales
  
  # En este punto nos concentramos solamente en las filas con productos
  # energéticos.
  
  oferta_monetaria <- oferta[c(38,39,43,46,98,99,100,
                     101,102,103,137,138,142),
                   -c(129, 130,135,147,148,149,
                      150,153,155,160,165,166)]
  
  # Creamos una versión solamente con la distribución porcentual de 
  # producción.
  
  distribucion_oferta <- oferta[c(38,39,43,46,98,99,100,
                                    101,102,103,137,138,142),
                                    -c(129,130,135,147:166)]
  
  # Llenamos con cero las celdas vacías por conveniencia
  distribucion_oferta[is.na(distribucion_oferta)] <- 0
  
  suma_distribucion_oferta <- rowSums(distribucion_oferta)
  # Ceros
  suma_distribucion_oferta[suma_distribucion_oferta == 0] <- 0.000001 
  
  # Obtenemos la distribución porcentual de lo monetario
  distribucion_oferta <- distribucion_oferta / suma_distribucion_oferta
  
  # Limpiamos
  rm(suma_distribucion_oferta)
  
  
  # Procesamiento del Cuadro de Utilización
  # =======================================
  
  id_cuadro  <- 4 # Utilización energética
  
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
  #   uc160 Subtotal Gobierno General
  #   uc161 Total gasto de consumo final
  #   uc165 P.5 TOTAL FORMACION BRUTA DE CAPITAL
  #   uc166	TOTAL UTILIZACIÓN
  
  #   Solamente nos concentramos en los productos energéticos en las filas.
  
  # A través de índices eliminamos lo que no se necesita.
  utilizacion_monetaria <- utilizacion[c(38,39,43,46,98,99,100,
                               101,102,103,137,138,142),
                             -c(129, 130, 135, 147, 148, 149, 150,
                                153, 157, 160, 161, 165, 166)]
  
  distribucion_utilizacion <- utilizacion[c(38,39,43,46,98,99,100,
                                            101,102,103,137,138,142),
                                          -c(129,130,135,147:153,
                                             157,160:166)]

  # Llenamos con cero las celdas vacías por conveniencia
  distribucion_utilizacion[is.na(distribucion_utilizacion)] <- 0
  
  # Obtenemos la suma de columnas que equivale a nuestro 100%
  suma_distribucion_utilizacion <- rowSums(distribucion_utilizacion)
  # Ceros
  suma_distribucion_utilizacion[suma_distribucion_utilizacion == 0] <- 0.000001 
  
  # Obtenemos la distribución porcentual de lo monetario
  distribucion_utilizacion <- distribucion_utilizacion / suma_distribucion_utilizacion
  
  # Y limpiamos
  rm( suma_distribucion_utilizacion)
  
  
  # Traemos los balances energéticos ordenados en forma de cuenta para
  # obtener los totales
  energiaTotal <- dbGetQuery(con,paste(
    'SELECT 
      id_ntgeM,
      id_npg4,
      sum(valor) AS "Tj" 
    FROM 
      balances_energeticos
    WHERE
      anio = ', 
    anio , 
    ' and id_npg4 != "" 
    GROUP BY 
      id_ntgeM,
      id_npg4
    ORDER BY
      id_ntgeM,
      id_npg4')
  )
  
  # id_ntgeM  ntgeM
  # 1 	      Producción
  # 2 	      Importaciones de bienes y servicios
  # 4 	      Consumo de energéticos (intermedio y final)
  # 6 	      Exportaciones de bienes y servicios
  # 7 	      Formación de capital
  # 8 	      Pérdidas
  # 9 	      No aprovechado
  
  # Lo ponemos en formato de matriz
  energiaTotal <- as.data.frame(tapply(energiaTotal$Tj, 
                         list(energiaTotal$id_npg4, 
                              energiaTotal$id_ntgeM), 
                         sum))

  # Y multiplicamos la distribucion_oferta y distribucion_utilizacion por
  # sus totales, para distribuirlo.
  distribucion_oferta2 <- distribucion_oferta * energiaTotal$`1`
  distribucion_util2 <- distribucion_utilizacion * energiaTotal$`4`
  
  # Y el remanente lo agregamos como una columna de "el resto".
  
  distribucion_oferta2 <- cbind(distribucion_oferta2,as.matrix(energiaTotal$`1`) - as.matrix(rowSums(distribucion_oferta2)))
  distribucion_util2 <- cbind(distribucion_util2,as.matrix(energiaTotal$`4`) - as.matrix(rowSums(distribucion_util2)))

  # Agregamos las columnas que nos hacen falta en la oferta y en la utilización
  
  ofertaE <- cbind(distribucion_oferta2, energiaTotal$`2`)
  utilizacionE <- cbind(distribucion_util2, energiaTotal$`6`, energiaTotal$`7`, energiaTotal$`8`, energiaTotal$`9`)
  
  # Agregamos los productos que no están en la parte monetaria.
  
  insumosAmbientales <- dbGetQuery(con,paste(
    'SELECT 
      id_ntgeM,
      id_energetico,
      sum(valor) AS "Tj" 
    FROM 
      balances_energeticos
    WHERE
      anio = ', 
    anio , 
    ' and id_energetico IN (\'EP04\', \'EP05\', \'EP09\') 
    GROUP BY 
      id_ntgeM,
      id_energetico
    ORDER BY
      id_ntgeM,
      id_energetico')
  )
  
  insumosAmbientales <- as.data.frame(tapply(insumosAmbientales$Tj, 
                                             list(insumosAmbientales$id_energetico, 
                                                  insumosAmbientales$id_ntgeM), 
                                             sum))
  
  # Agregamos los insumos ambientales a la oferta
  
  ofertaE <- cbind(ofertaE, 0)
  ofertaE <- rbind(matrix(0,nrow= dim(insumosAmbientales[1]), ncol = dim(ofertaE)[2]),ofertaE)
  ofertaE[c(1:dim(insumosAmbientales)[1]), dim(ofertaE)[2]] <- insumosAmbientales$`1`
  
  # Puesto que los insumos ambientales se utilizan por las actividades
  # no se agrega columna de estos al cuadro de utilización, pero sí 
  # se agregan las filas correspondientes y los valores a la actividad que los usa
  utilizacionE <- rbind(matrix(0,nrow= dim(insumosAmbientales[1]), ncol = dim(utilizacionE)[2]),utilizacionE)
  utilizacionE[c(1:dim(insumosAmbientales)[1]), match("GTMuc082",colnames(utilizacionE))] <- insumosAmbientales$`4`

  # Y renombramos los elementos recién agregados en cada cuadro
  colnames(ofertaE)[c((length(colnames(ofertaE))-2) : length(colnames(ofertaE)))] <- c("GTMoc000a", "GTMoc151","GTMoc000b")
  colnames(utilizacionE)[c((length(colnames(utilizacionE))-4) : length(colnames(utilizacionE)))] <- c("GTMuc000a", "GTMuc151","GTMuc163","GTMuc000b","GTMuc000c")
  
  rownames(ofertaE)[ c(1: dim(insumosAmbientales)[1])] <- c("GTMf000a","GTMf000b", "GTMf000c")
  rownames(utilizacionE)[ c(1: dim(insumosAmbientales)[1])] <- c("GTMf000a","GTMf000b", "GTMf000c")
  
  # ========================
  # Emisiones CO2, CH4 y N2o 
  # ========================
  
  # Multiplicacion por coeficientes elemento por elemento
  # (No es multiplicación de matrices)
  
  emisiones_CO2 <- utilizacionE * CO2_coef / 1000
  emisiones_CH4 <- utilizacionE * CH4_coef / 1000
  emisiones_N2O <- utilizacionE * N2O_coef / 1000
  
  # Melts de cuadros procesados
  # Melt 
  
  # Cuadros: 
  # 1 Oferta
  # 2 Utilización
  # 3 Oferta energética monetaria
  # 4 Utilización energética monetaria
  # 5 Oferta energética física
  # 6 Utilización energética física
  # 7 Oferta de CO2 equivalente
  # 8 Oferta de CH4 equivalente
  # 9 Oferta de N2O equivalente
  
  # Unidad:
  # 1 Millones de quetzales
  # 2 Millones de quetzales en medidas encadenadas 
  #   de volumen con año de referencia 2013
  # 3 Terajulios
  # 4 Toneladas métricas
  # 5 Toneladas equivalentes de dióxido de carbono
  
  oferta_energetica_monetaria <- cbind(iso3, 
                                       anio,
                                       3, 
                                       melt(oferta_monetaria), 
                                       1)
  colnames(oferta_energetica_monetaria) <-
    c("iso3",
      "anio",
      "id_cuadro",
      "id_fila",
      "id_columna",
      "valor",
      "id_unidad")

  utilizacion_energetica_monetaria <- cbind(iso3, 
                                       anio,
                                       4, 
                                       melt(utilizacion_monetaria), 
                                       1)
  colnames(utilizacion_energetica_monetaria) <-
    c("iso3",
      "anio",
      "id_cuadro",
      "id_fila",
      "id_columna",
      "valor",
      "id_unidad")
  
  oferta_energetica_fisica <- cbind(iso3, 
                                       anio,
                                       5, 
                                       melt(ofertaE), 
                                       3)
  colnames(oferta_energetica_fisica) <-
    c("iso3",
      "anio",
      "id_cuadro",
      "id_fila",
      "id_columna",
      "valor",
      "id_unidad")
  
  utilizacion_energetica_fisica <- cbind(iso3, 
                                            anio,
                                            6, 
                                            melt(utilizacionE), 
                                            3)
  colnames(utilizacion_energetica_fisica) <-
    c("iso3",
      "anio",
      "id_cuadro",
      "id_fila",
      "id_columna",
      "valor",
      "id_unidad")
  
  oferta_gei_co2_eq <- cbind(iso3, 
                                    anio,
                                    7, 
                                    melt(emisiones_CO2), 
                                    5)
  colnames(oferta_gei_co2_eq) <-
    c("iso3",
      "anio",
      "id_cuadro",
      "id_fila",
      "id_columna",
      "valor",
      "id_unidad")
  
  oferta_gei_ch4_eq <- cbind(iso3, 
                             anio,
                             8, 
                             melt(emisiones_CH4*84), # CH4 GWP 20 años 
                             5)
  colnames(oferta_gei_ch4_eq) <-
    c("iso3",
      "anio",
      "id_cuadro",
      "id_fila",
      "id_columna",
      "valor",
      "id_unidad")
  
  oferta_gei_n2o_eq <- cbind(iso3, 
                             anio,
                             9, 
                             melt(emisiones_N2O*298), # N2O GWP 20 años 
                             5)
  colnames(oferta_gei_n2o_eq) <-
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
    union <- rbind(oferta_energetica_monetaria,
                   utilizacion_energetica_monetaria,
                   oferta_energetica_fisica,
                   utilizacion_energetica_fisica,
                   oferta_gei_co2_eq,
                   oferta_gei_ch4_eq,
                   oferta_gei_n2o_eq
    )
    
    assign(paste("COU_E_", anio, sep = ""), union)
  } else {
    union <- rbind(oferta_energetica_monetaria,
                   utilizacion_energetica_monetaria,
                   oferta_energetica_fisica,
                   utilizacion_energetica_fisica,
                   oferta_gei_co2_eq,
                   oferta_gei_ch4_eq,
                   oferta_gei_n2o_eq
    )
    assign(paste("COU_E_", anio, sep = ""), union)
  }
  lista <- c(lista, paste("COU_E_", anio, sep = ""))
}



# Eliminamos el primer elemento provisional que usamos para inicializar
# la lista, pues R no nos permite declarar listas vacías. 
lista <- lapply(lista[-1], as.name)

# Unimos los objetos de todos los años y precios
SCN_E <- do.call(rbind.data.frame, lista)

# Por conveniencia, convertimos los valores NA a 0
SCN_E$valor[is.na(SCN_E$valor)] <- 0

# Y borramos los objetos individuales
do.call(rm,lista)

# Recolección de basura
gc()

head(SCN_E)

# Base de datos

dbSendStatement(con, "DROP TABLE IF EXISTS scn_e")
dbCreateTable(con, "scn_e", SCN_E)
dbAppendTable(con, "scn_e", SCN_E)

# Crear vista
query <- read_file("scripts/VIEW_cuenta_energia.sql")
dbSendStatement(con, "DROP VIEW IF EXISTS cuenta_energia")
dbSendQuery(con, query)

# Leer vista para Excel
SCN_E2 <- dbGetQuery(con,"
SELECT 
  * 
FROM 
  cuenta_energia 
")


# Cerramos la conexión a la base de datos
dbDisconnect(con)

write.xlsx(
  SCN_E2,
  "salidas/SCN_E.xlsx",
  sheetName= "SCNE_BD",
  rowNames=FALSE,
  colnames=FALSE,
  overwrite = TRUE,
  asTable = FALSE
)

