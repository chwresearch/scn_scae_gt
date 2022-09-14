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
  unidad <- toString(info[1,1 ]) # unidad de medida
  # precios Corrientes o medidas Encadenados
  if (unidad != "Millones de quetzales") {
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
  rownames(oferta) <- c(sprintf("of%03d", seq(1, dim(oferta)[1])))
  colnames(oferta) <- c(sprintf("oc%03d", seq(1, dim(oferta)[2])))
  
  # Columnas a eliminar con subtotales y totales
  
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
  
  oferta <- oferta[-c(214, 215, 216, 217, 218, 219, 224, 225), -c(130,135,147,148,149,150,153,155,160,165,166)]
  
  # Desdoblamos
  oferta <- cbind(anio, id_precios,1, melt(oferta), id_unidad)
  
  colnames(oferta) <-
    c("Año",
      "Id. Precios",
      "Id. Cuadro",
      "Filas",
      "Columnas",
      "Valor",
      "Id. Unidad")
  
  # Cuadro de utilización
  # =====================
  
  utilizacion <- as.matrix(read_excel(
    archivo,
    range = paste("'" , hojas[i], "'!C252:FL476", sep = ""),
    col_names = FALSE,
    col_types = "numeric"
  ))
  rownames(utilizacion) <-
    c(sprintf("uf%03d", seq(1, dim(utilizacion)[1])))
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
  #   uc155 Desagregación irrelevante
  #   uc156 Desagregación irrelevante
  #   uc160 Subtotal Gobienro General
  #   uc161 Total gasto de consumo final
  #   uc165 P.5 TOTAL FORMACION BRUTA DE CAPITAL
  #   uc166	TOTAL UTILIZACIÓN
  
  #   Filas a eliminar
  #   214, 215, 216, 218, 219, 224, 225
  
  utilizacion <- utilizacion[-c(214, 215, 216, 218, 219, 224, 225),
                             -c(130, 135, 147, 148, 149, 150, 153,
                                155, 156, 160, 161, 165, 166)]
  
  # Desdoblamos
  utilizacion <-
    cbind(anio, id_precios,1, melt(utilizacion), id_unidad)
  
  colnames(utilizacion) <-
    c("Año",
      "Id. Precios",
      "Id. Cuadro",
      "Filas",
      "Columnas",
      "Valor",
      "Id. Unidad")
  
  
}
