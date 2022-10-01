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
  
  id_cuadro <- 3 # 1  Oferta energética
  
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
  rm(oferta, suma_distribucion_oferta)
  
  # Melt 
    
  # oferta <- cbind(iso3, anio,id_cuadro, melt(oferta), id_unidad)
  # colnames(oferta) <-
  #   c("iso3",
  #     "anio",
  #     "id_cuadro",
  #     "id_fila",
  #     "id_columna",
  #     "valor",
  #     "id_unidad")
  
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
  rm(utilizacion, suma_distribucion_utilizacion)
  
  # utilizacion <- cbind(iso3,anio, id_cuadro, melt(utilizacion), id_unidad)
  # 
  # colnames(utilizacion) <-
  #   c("iso3",
  #     "anio",
  #     "id_cuadro",
  #     "id_fila",
  #     "id_columna",
  #     "valor",
  #     "id_unidad")
  
  
  # Traemos la producción y consumo intermedio y final energéticos
  
  con <- dbConnect(RSQLite::SQLite(), "datos/scn.db")
  prod_cons_energetico <- 
  join(
    dbGetQuery(con,paste(
    'SELECT 
      id_npg4,
      sum(valor) AS "Producción" 
    FROM 
      balances_energeticos
    WHERE
      anio = ', 
    anio , 
    ' and id_ntgeM = 1
      and id_npg4 != "" 
    GROUP BY 
      id_npg4
    ORDER BY
      id_npg4')
    )
    ,
    dbGetQuery(con,paste(
      'SELECT 
      id_npg4,
      sum(valor) AS "Consumo intermedio y final"
    FROM 
      balances_energeticos
    WHERE
      anio = ', 
      anio , 
      ' and id_ntgeM = 4
      and id_npg4 != "" 
    GROUP BY 
      id_npg4
    ORDER BY
      id_npg4')
    )
    ,
    by = "id_npg4"
  )  
  
  dbDisconnect(con)
  
  # Y multiplicamos la distribucion_oferta y distribucion_utilizacion por
  # sus totales, para distribuirlo.
  distribucion_oferta2 <- distribucion_oferta * prod_cons_energetico$`Producción`
  distribucion_util2 <- distribucion_utilizacion * prod_cons_energetico$`Consumo intermedio y final`
  
  distribucion_oferta2 <- cbind(distribucion_oferta2,as.matrix(prod_cons_energetico$`Producción`) - as.matrix(rowSums(distribucion_oferta2)))
  distribucion_util2 <- cbind(distribucion_util2,as.matrix(prod_cons_energetico$`Consumo intermedio y final`) - as.matrix(rowSums(distribucion_util2)))
  
  # Renombramos la columna recién creada en cada cuadro
  
  colnames(distribucion_oferta2)[length(colnames(distribucion_oferta2))] <- paste(iso3,"oc", length(colnames(distribucion_oferta2)) , sep = "")
  
  
  
  # Unión de los cuadros procesados
  # ===============================
  
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



