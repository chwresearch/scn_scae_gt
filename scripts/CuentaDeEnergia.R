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
library(stringr)
library(RSQLite)
library(DBI)

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
  
  # Por el momento no se incluirá precios constantes.
  
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


